# Azure Deployment - Quick Start

Deploy Sigma Thermal Calculators to **Azure Static Web Apps** in **East US 2**

---

## ⚡ Quick Deploy (5 minutes)

### Prerequisites

- Azure account: https://azure.microsoft.com/free/
- Azure CLI: https://aka.ms/azure-cli
- GitHub account

### One-Command Deployment

```bash
./deploy-azure.sh
```

This script will:
1. Login to Azure
2. Create resource group in East US 2
3. Create Static Web App
4. Connect to GitHub
5. Provide deployment token

---

## 📋 Manual Steps (Alternative)

### 1. Install Azure CLI

```bash
# macOS
brew install azure-cli

# Windows
winget install Microsoft.AzureCLI

# Linux
curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash
```

### 2. Deploy

```bash
# Login
az login

# Create resources
az group create --name sigma-thermal-rg --location eastus2

az staticwebapp create \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --location eastus2 \
  --source https://github.com/YOUR_USERNAME/sigma-thermal \
  --branch main \
  --app-location "/web" \
  --api-location "/api" \
  --login-with-github
```

### 3. Get Deployment Token

```bash
az staticwebapp secrets list \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --query "properties.apiKey" -o tsv
```

### 4. Add GitHub Secret

1. Go to: https://github.com/YOUR_USERNAME/sigma-thermal/settings/secrets/actions
2. Click "New repository secret"
3. Name: `AZURE_STATIC_WEB_APPS_API_TOKEN`
4. Value: Paste token from step 3
5. Click "Add secret"

### 5. Deploy Code

```bash
git add .
git commit -m "Deploy to Azure"
git push origin main
```

---

## 🌐 Your App URL

After deployment (2-3 minutes):

```bash
az staticwebapp show \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --query "defaultHostname" -o tsv
```

Access at: `https://[generated-name].azurestaticapps.net`

---

## ✅ Test Deployment

### Test Static Pages

```bash
curl https://YOUR-APP-URL.azurestaticapps.net/
curl https://YOUR-APP-URL.azurestaticapps.net/resource.html
```

### Test API

```bash
curl -X POST https://YOUR-APP-URL.azurestaticapps.net/api/calculate/heating-value \
  -H "Content-Type: application/json" \
  -d '{"ch4":100,"c2h6":0,"c3h8":0,"c4h10":0,"h2":0,"co":0,"h2s":0,"co2":0,"n2":0}'
```

Expected: JSON response with heating values

---

## 💰 Cost

**Free Tier** (recommended):
- 100 GB bandwidth/month
- SSL certificates included
- Custom domains included
- **Cost: $0/month**

---

## 📁 What Was Created

```
✅ web/staticwebapp.config.json    - Azure SWA configuration
✅ api/host.json                    - Functions runtime config
✅ api/heating_value/__init__.py    - Heating value API endpoint
✅ api/heating_value/function.json  - Function binding config
✅ api/requirements.txt             - Python dependencies
✅ .github/workflows/azure-static-web-apps.yml - CI/CD pipeline
✅ AZURE_DEPLOYMENT.md              - Full documentation
✅ deploy-azure.sh                  - Quick deployment script
```

---

## 🔧 Troubleshooting

### Issue: "az: command not found"
**Solution:** Install Azure CLI: https://aka.ms/azure-cli

### Issue: GitHub authentication fails
**Solution:** Run `az login` again and authorize GitHub access

### Issue: API returns 500 error
**Solution:** Check Application Insights logs or run locally:
```bash
cd api
func start
```

---

## 📖 Full Documentation

For complete details, see: `AZURE_DEPLOYMENT.md`

---

## 🎯 Next Steps

1. ✅ Deploy to Azure (this guide)
2. Test all calculator endpoints
3. Add custom domain (optional)
4. Set up monitoring with Application Insights
5. Create remaining API endpoints:
   - Air Requirement Calculator
   - Products of Combustion
   - Steam Properties
   - Water Properties
   - Flash Steam

---

**Region:** East US 2
**Deployment Time:** ~3 minutes
**Cost:** Free (Free tier)
**Status:** Ready to deploy
