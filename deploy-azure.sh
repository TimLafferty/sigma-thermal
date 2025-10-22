#!/bin/bash
# Quick Azure Deployment Script for Sigma Thermal Calculators
# Region: East US 2

set -e

echo "🚀 Sigma Thermal Calculators - Azure Deployment"
echo "================================================"
echo ""

# Configuration
RESOURCE_GROUP="sigma-thermal-rg"
APP_NAME="sigma-thermal-calculators"
LOCATION="eastus2"
REPO_URL="https://github.com/YOUR_USERNAME/sigma-thermal"
BRANCH="main"

# Check if Azure CLI is installed
if ! command -v az &> /dev/null; then
    echo "❌ Azure CLI is not installed"
    echo "Install from: https://aka.ms/azure-cli"
    exit 1
fi

echo "✅ Azure CLI found"

# Login to Azure
echo ""
echo "🔐 Logging in to Azure..."
az login

# Select subscription
echo ""
echo "📋 Available subscriptions:"
az account list --output table

echo ""
read -p "Enter subscription ID or name (or press Enter for default): " SUBSCRIPTION
if [ -n "$SUBSCRIPTION" ]; then
    az account set --subscription "$SUBSCRIPTION"
fi

CURRENT_SUBSCRIPTION=$(az account show --query name -o tsv)
echo "✅ Using subscription: $CURRENT_SUBSCRIPTION"

# Create resource group
echo ""
echo "📦 Creating resource group: $RESOURCE_GROUP in $LOCATION"
az group create \
    --name $RESOURCE_GROUP \
    --location $LOCATION \
    --output table

# Get GitHub username
echo ""
read -p "Enter your GitHub username: " GITHUB_USERNAME
REPO_URL="https://github.com/$GITHUB_USERNAME/sigma-thermal"

# Create Static Web App
echo ""
echo "🌐 Creating Azure Static Web App: $APP_NAME"
echo "   This will open GitHub authorization in your browser..."
echo ""

az staticwebapp create \
    --name $APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --location $LOCATION \
    --source $REPO_URL \
    --branch $BRANCH \
    --app-location "/web" \
    --api-location "/api" \
    --output-location "" \
    --login-with-github

# Get deployment token
echo ""
echo "🔑 Getting deployment token..."
DEPLOYMENT_TOKEN=$(az staticwebapp secrets list \
    --name $APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --query "properties.apiKey" -o tsv)

# Get app URL
APP_URL=$(az staticwebapp show \
    --name $APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --query "defaultHostname" -o tsv)

# Display results
echo ""
echo "========================================"
echo "✅ Deployment Complete!"
echo "========================================"
echo ""
echo "📱 Your App URL:"
echo "   https://$APP_URL"
echo ""
echo "🔑 GitHub Secret (copy this):"
echo "   Name:  AZURE_STATIC_WEB_APPS_API_TOKEN"
echo "   Value: $DEPLOYMENT_TOKEN"
echo ""
echo "📝 Next Steps:"
echo "   1. Go to: https://github.com/$GITHUB_USERNAME/sigma-thermal/settings/secrets/actions"
echo "   2. Click 'New repository secret'"
echo "   3. Name: AZURE_STATIC_WEB_APPS_API_TOKEN"
echo "   4. Paste the value above"
echo "   5. Push to main branch to trigger deployment"
echo ""
echo "   git add ."
echo "   git commit -m 'Add Azure deployment'"
echo "   git push origin main"
echo ""
echo "⏱️  Initial deployment takes 2-3 minutes"
echo "📊 Monitor at: https://github.com/$GITHUB_USERNAME/sigma-thermal/actions"
echo ""
echo "========================================"
