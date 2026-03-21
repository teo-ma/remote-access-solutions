#!/bin/bash
#==============================================================================
# AVD Private Link Demo - Azure China (21Vianet)
#
# 架构:
#   ┌─────────────┐  Private Link   ┌─────────────┐  Private VNet  ┌──────────┐
#   │ Client VM   │ ──────────────> │ AVD Service  │ ────────────> │ Session  │
#   │ (Win11 Ent) │   feed+conn     │ (HostPool+WS)│               │ Host VM  │
#   │ 10.0.1.x    │                 │              │               │ (Win11)  │
#   └─────────────┘                 └──────────────┘               │ 10.0.2.x │
#         │                                                        └────┬─────┘
#    Azure Bastion                                                      │
#    (管理访问)                                                    Private VNet
#                                                                       │
#                                                               ┌───────▼──────┐
#                                                               │ Backend Linux│
#                                                               │  (nginx)     │
#                                                               │  10.0.3.x    │
#                                                               └──────────────┘
#
# 网络说明:
#   - 所有 VM 无公网 IP，通过 NAT Gateway 出站
#   - AVD 客户端通过 Private Link 连接 AVD 服务
#   - Session Host 通过 VNet 内网连接 Backend Linux
#   - 管理访问通过 Azure Bastion
#==============================================================================
set -euo pipefail

# ========================= 变量 ==============================================
RG="rg-avd-privatelink-demo"
LOCATION="chinanorth3"
SUB_ID=$(az account show --query id -o tsv)

# 网络
VNET_NAME="vnet-avd-demo"
VNET_PREFIX="10.0.0.0/16"
CLIENT_SUBNET="client-subnet"
CLIENT_PREFIX="10.0.1.0/24"
AVD_SUBNET="avd-subnet"
AVD_PREFIX="10.0.2.0/24"
BACKEND_SUBNET="backend-subnet"
BACKEND_PREFIX="10.0.3.0/24"
PE_SUBNET="pe-subnet"
PE_PREFIX="10.0.4.0/24"
BASTION_PREFIX="10.0.5.0/26"

# AVD
HOSTPOOL="hp-avd-demo"
WORKSPACE="ws-avd-demo"
APPGROUP="ag-avd-demo"

# VM
ADMIN_USER="avdadmin"
ADMIN_PASS='AVDdemo#2024Secure'   # 请按需修改
SESSION_HOST="vm-avd-host"
CLIENT_VM="vm-avd-client"
LINUX_VM="vm-backend"

WIN_SIZE="Standard_D2s_v5"
LINUX_SIZE="Standard_B2s"

echo "============================================"
echo " AVD Private Link Demo 部署开始"
echo " 订阅: $SUB_ID"
echo " 区域: $LOCATION"
echo "============================================"

#==============================================================================
# STEP 1: 资源组 & Provider 注册
#==============================================================================
echo ""
echo ">>> [1/9] 创建资源组..."
az group create -n "$RG" -l "$LOCATION" -o none
az provider register -n Microsoft.DesktopVirtualization -o none 2>/dev/null || true

#==============================================================================
# STEP 2: 网络基础设施 (VNet / Subnets / NAT GW / NSG)
#==============================================================================
echo ">>> [2/9] 创建网络基础设施..."

# VNet
az network vnet create -g "$RG" -n "$VNET_NAME" \
  --address-prefix "$VNET_PREFIX" -l "$LOCATION" -o none

# 子网
az network vnet subnet create -g "$RG" --vnet-name "$VNET_NAME" \
  -n "$CLIENT_SUBNET" --address-prefix "$CLIENT_PREFIX" -o none
az network vnet subnet create -g "$RG" --vnet-name "$VNET_NAME" \
  -n "$AVD_SUBNET" --address-prefix "$AVD_PREFIX" -o none
az network vnet subnet create -g "$RG" --vnet-name "$VNET_NAME" \
  -n "$BACKEND_SUBNET" --address-prefix "$BACKEND_PREFIX" -o none
az network vnet subnet create -g "$RG" --vnet-name "$VNET_NAME" \
  -n "$PE_SUBNET" --address-prefix "$PE_PREFIX" -o none
az network vnet subnet create -g "$RG" --vnet-name "$VNET_NAME" \
  -n AzureBastionSubnet --address-prefix "$BASTION_PREFIX" -o none

# 禁用 PE 子网的 private endpoint 网络策略
az network vnet subnet update -g "$RG" --vnet-name "$VNET_NAME" \
  -n "$PE_SUBNET" --private-endpoint-network-policies Disabled -o none

# NAT Gateway (所有 VM 出站互联网)
echo "    创建 NAT Gateway..."
az network public-ip create -g "$RG" -n pip-natgw --sku Standard -l "$LOCATION" -o none
az network nat gateway create -g "$RG" -n natgw-avd \
  --public-ip-addresses pip-natgw -l "$LOCATION" --idle-timeout 10 -o none
for SUBNET in "$CLIENT_SUBNET" "$AVD_SUBNET" "$BACKEND_SUBNET"; do
  az network vnet subnet update -g "$RG" --vnet-name "$VNET_NAME" \
    -n "$SUBNET" --nat-gateway natgw-avd -o none
done

# NSG
echo "    创建 NSG..."
az network nsg create -g "$RG" -n nsg-client -l "$LOCATION" -o none
az network nsg create -g "$RG" -n nsg-avd -l "$LOCATION" -o none
az network nsg create -g "$RG" -n nsg-backend -l "$LOCATION" -o none

# Backend NSG: 仅允许 AVD 子网 HTTP/HTTPS
az network nsg rule create -g "$RG" --nsg-name nsg-backend \
  -n AllowHTTPFromAVD --priority 1000 --direction Inbound \
  --access Allow --protocol Tcp --destination-port-ranges 80 443 \
  --source-address-prefixes "$AVD_PREFIX" -o none

# 关联 NSG
az network vnet subnet update -g "$RG" --vnet-name "$VNET_NAME" \
  -n "$CLIENT_SUBNET" --nsg nsg-client -o none
az network vnet subnet update -g "$RG" --vnet-name "$VNET_NAME" \
  -n "$AVD_SUBNET" --nsg nsg-avd -o none
az network vnet subnet update -g "$RG" --vnet-name "$VNET_NAME" \
  -n "$BACKEND_SUBNET" --nsg nsg-backend -o none

#==============================================================================
# STEP 3: Azure Bastion (管理访问，创建耗时约 10-20 分钟)
#==============================================================================
echo ">>> [3/9] 创建 Azure Bastion (耗时 ~15 分钟)..."
az network public-ip create -g "$RG" -n pip-bastion --sku Standard -l "$LOCATION" -o none
az network bastion create -g "$RG" -n bastion-avd -l "$LOCATION" \
  --vnet-name "$VNET_NAME" --public-ip-address pip-bastion --sku Basic --no-wait -o none
echo "    Bastion 已开始创建 (--no-wait), 后续步骤继续..."

#==============================================================================
# STEP 4: AVD 服务组件 (Host Pool / App Group / Workspace)
#==============================================================================
echo ">>> [4/9] 创建 AVD 服务组件..."

# 生成 24 小时后过期的注册令牌时间
EXPIRATION=$(date -u -v+24H '+%Y-%m-%dT%H:%M:%S.0000000Z')

# 创建 Host Pool (AVD 控制平面只在 chinanorth2/chinaeast2 可用)
AVD_LOCATION="chinanorth2"
az desktopvirtualization hostpool create \
  -g "$RG" -n "$HOSTPOOL" -l "$AVD_LOCATION" \
  --host-pool-type Pooled \
  --load-balancer-type BreadthFirst \
  --preferred-app-group-type Desktop \
  --max-session-limit 10 \
  --registration-info registration-token-operation="Update" expiration-time="$EXPIRATION" \
  -o none

# 创建 Desktop Application Group
HP_ID="/subscriptions/$SUB_ID/resourceGroups/$RG/providers/Microsoft.DesktopVirtualization/hostPools/$HOSTPOOL"
az desktopvirtualization applicationgroup create \
  -g "$RG" -n "$APPGROUP" -l "$AVD_LOCATION" \
  --application-group-type Desktop \
  --host-pool-arm-path "$HP_ID" \
  -o none

# 创建 Workspace 并关联 App Group
AG_ID="/subscriptions/$SUB_ID/resourceGroups/$RG/providers/Microsoft.DesktopVirtualization/applicationGroups/$APPGROUP"
az desktopvirtualization workspace create \
  -g "$RG" -n "$WORKSPACE" -l "$AVD_LOCATION" \
  --application-group-references "$AG_ID" \
  -o none

# 获取注册令牌
echo "    获取注册令牌..."
TOKEN=$(az desktopvirtualization hostpool retrieve-registration-token \
  -g "$RG" -n "$HOSTPOOL" --query token -o tsv 2>/dev/null) || {
  echo "    CLI retrieve-registration-token 不可用，使用 REST API..."
  TOKEN=$(az rest --method POST \
    --url "https://management.chinacloudapi.cn/subscriptions/$SUB_ID/resourceGroups/$RG/providers/Microsoft.DesktopVirtualization/hostPools/$HOSTPOOL/retrieveRegistrationToken?api-version=2024-04-03" \
    --body '{}' --query token -o tsv)
}
echo "    注册令牌已获取: ${TOKEN:0:20}..."

#==============================================================================
# STEP 5: AVD Session Host VM (Windows 11 多会话 + AVD Agent + AAD Join)
#==============================================================================
echo ">>> [5/9] 创建 AVD Session Host VM..."
az vm create -g "$RG" -n "$SESSION_HOST" -l "$LOCATION" \
  --image MicrosoftWindowsDesktop:windows-11:win11-24h2-avd:latest \
  --size "$WIN_SIZE" \
  --admin-username "$ADMIN_USER" --admin-password "$ADMIN_PASS" \
  --vnet-name "$VNET_NAME" --subnet "$AVD_SUBNET" \
  --public-ip-address "" --nsg "" \
  --license-type Windows_Client \
  -o none

echo "    安装 AVD Agent (DSC 扩展)..."
az vm extension set -g "$RG" --vm-name "$SESSION_HOST" \
  --name DSC --publisher Microsoft.Powershell --version 2.83 \
  --settings "{
    \"modulesUrl\": \"https://wvdportalstorageblob.blob.core.windows.net/galleryartifacts/Configuration_09-08-2022.zip\",
    \"configurationFunction\": \"Configuration.ps1\\\\AddSessionHost\",
    \"properties\": {
      \"HostPoolName\": \"$HOSTPOOL\",
      \"RegistrationInfoToken\": \"$TOKEN\",
      \"aadJoin\": true
    }
  }" -o none

echo "    配置 Entra ID (AAD) 加入..."
az vm extension set -g "$RG" --vm-name "$SESSION_HOST" \
  --name AADLoginForWindows \
  --publisher Microsoft.Azure.ActiveDirectory \
  -o none

#==============================================================================
# STEP 6: Client VM (Windows 11 企业版 - 模拟 AVD 客户端)
#==============================================================================
echo ">>> [6/9] 创建 Client VM (Windows 11)..."
az vm create -g "$RG" -n "$CLIENT_VM" -l "$LOCATION" \
  --image MicrosoftWindowsDesktop:windows-11:win11-24h2-ent:latest \
  --size "$WIN_SIZE" \
  --admin-username "$ADMIN_USER" --admin-password "$ADMIN_PASS" \
  --vnet-name "$VNET_NAME" --subnet "$CLIENT_SUBNET" \
  --public-ip-address "" --nsg "" \
  --license-type Windows_Client \
  -o none

#==============================================================================
# STEP 7: Backend Linux VM + nginx (模拟应用服务器)
#==============================================================================
echo ">>> [7/9] 创建 Backend Linux VM + nginx..."
az vm create -g "$RG" -n "$LINUX_VM" -l "$LOCATION" \
  --image Canonical:0001-com-ubuntu-server-jammy:22_04-lts:latest \
  --size "$LINUX_SIZE" \
  --admin-username "$ADMIN_USER" --admin-password "$ADMIN_PASS" \
  --authentication-type password \
  --vnet-name "$VNET_NAME" --subnet "$BACKEND_SUBNET" \
  --public-ip-address "" --nsg "" \
  -o none

echo "    安装 nginx..."
az vm extension set -g "$RG" --vm-name "$LINUX_VM" \
  --name customScript --publisher Microsoft.Azure.Extensions \
  --settings '{"commandToExecute":"apt-get update && apt-get install -y nginx && systemctl enable nginx && systemctl start nginx && echo \"<h1>AVD Private Link Demo - Backend App Server</h1><p>Hostname: $(hostname)</p><p>IP: $(hostname -I)</p>\" > /var/www/html/index.html"}' \
  -o none

#==============================================================================
# STEP 8: AVD Private Link (Private Endpoints + DNS)
#==============================================================================
echo ">>> [8/9] 配置 AVD Private Link..."

HP_RESOURCE="/subscriptions/$SUB_ID/resourceGroups/$RG/providers/Microsoft.DesktopVirtualization/hostPools/$HOSTPOOL"
WS_RESOURCE="/subscriptions/$SUB_ID/resourceGroups/$RG/providers/Microsoft.DesktopVirtualization/workspaces/$WORKSPACE"

# Host Pool Private Endpoint (connection)
echo "    创建 Host Pool Private Endpoint..."
az network private-endpoint create -g "$RG" -n pe-hostpool \
  -l "$LOCATION" --vnet-name "$VNET_NAME" --subnet "$PE_SUBNET" \
  --private-connection-resource-id "$HP_RESOURCE" \
  --group-id connection \
  --connection-name pec-hostpool -o none

# Workspace Private Endpoint (feed)
echo "    创建 Workspace Private Endpoint..."
az network private-endpoint create -g "$RG" -n pe-workspace \
  -l "$LOCATION" --vnet-name "$VNET_NAME" --subnet "$PE_SUBNET" \
  --private-connection-resource-id "$WS_RESOURCE" \
  --group-id feed \
  --connection-name pec-workspace -o none

# Private DNS Zone
echo "    创建 Private DNS Zone..."
az network private-dns zone create -g "$RG" \
  -n privatelink.wvd.azure.cn -o none

az network private-dns link vnet create -g "$RG" \
  --zone-name privatelink.wvd.azure.cn \
  -n link-avd-vnet --virtual-network "$VNET_NAME" \
  --registration-enabled false -o none

# DNS Zone Groups (自动注册 DNS 记录)
az network private-endpoint dns-zone-group create -g "$RG" \
  --endpoint-name pe-hostpool --name dnsgroup-hp \
  --private-dns-zone privatelink.wvd.azure.cn \
  --zone-name wvd -o none

az network private-endpoint dns-zone-group create -g "$RG" \
  --endpoint-name pe-workspace --name dnsgroup-ws \
  --private-dns-zone privatelink.wvd.azure.cn \
  --zone-name wvd -o none

#==============================================================================
# STEP 9: 角色分配 & 验证
#==============================================================================
echo ">>> [9/9] 配置角色分配..."

UPN=$(az account show --query user.name -o tsv)

# 允许当前用户通过 AAD 登录 VM (用于 Bastion 连接)
az role assignment create --assignee "$UPN" \
  --role "Virtual Machine User Login" \
  --scope "/subscriptions/$SUB_ID/resourceGroups/$RG" -o none 2>/dev/null || true

# 允许当前用户访问 AVD 桌面
az role assignment create --assignee "$UPN" \
  --role "Desktop Virtualization User" \
  --scope "$AG_ID" -o none 2>/dev/null || true

#==============================================================================
# 输出摘要
#==============================================================================
echo ""
echo "=============================================="
echo "  ✅ AVD Private Link Demo 部署完成!"
echo "=============================================="
echo ""
echo "  资源组:        $RG"
echo "  Host Pool:     $HOSTPOOL"
echo "  Workspace:     $WORKSPACE"
echo "  App Group:     $APPGROUP"
echo ""

# 获取 VM 私有 IP
SH_IP=$(az vm show -g "$RG" -n "$SESSION_HOST" -d --query privateIps -o tsv 2>/dev/null || echo "pending")
CL_IP=$(az vm show -g "$RG" -n "$CLIENT_VM" -d --query privateIps -o tsv 2>/dev/null || echo "pending")
LX_IP=$(az vm show -g "$RG" -n "$LINUX_VM" -d --query privateIps -o tsv 2>/dev/null || echo "pending")

echo "  Session Host:  $SESSION_HOST ($SH_IP) - AVD subnet"
echo "  Client VM:     $CLIENT_VM ($CL_IP) - Client subnet"
echo "  Backend Linux: $LINUX_VM ($LX_IP) - Backend subnet"
echo ""
echo "  管理员用户:    $ADMIN_USER"
echo "  管理员密码:    $ADMIN_PASS"
echo ""
echo "  ── 下一步 ──"
echo "  1. 通过 Azure Bastion 连接 Client VM ($CLIENT_VM)"
echo "  2. 在 Client VM 上安装 Windows App (AVD 客户端)"
echo "  3. 登录 AVD workspace 验证 Private Link 连接"
echo "  4. 在 Session Host 中访问 http://$LX_IP 验证后端连接"
echo ""
echo "  ── 可选：禁用公共访问 ──"
echo "  az desktopvirtualization hostpool update -g $RG -n $HOSTPOOL --public-network-access Disabled"
echo "  az desktopvirtualization workspace update -g $RG -n $WORKSPACE --public-network-access Disabled"
echo "=============================================="
