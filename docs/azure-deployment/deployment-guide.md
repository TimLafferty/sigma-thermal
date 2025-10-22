# Azure Deployment Guide - Sigma Thermal Calculators

Deploy the Sigma Thermal Engineering Calculators to Azure Static Web Apps with Azure Functions backend.

**Region:** East US 2
**Date:** October 22, 2025

---

## Architecture

```
┌─────────────────────────────────────────┐
│   Azure Static Web Apps                 │
│   (East US 2)                           │
│                                         │
│  ┌──────────────────────────────────┐  │
│  │  Static Content (HTML/CSS/JS)    │  │
│  │  - index.html                    │  │
│  │  - resource.html                 │  │
│  │  - calculators/*.html            │  │
│  │  - css/style.css                 │  │
│  │  - js/*.js                       │  │
│  └──────────────────────────────────┘  │
│                                         │
│  ┌──────────────────────────────────┐  │
│  │  Azure Functions API             │  │
│  │  - /api/calculate/heating-value  │  │
│  │  - /api/calculate/air-requirement│  │
│  │  - /api/calculate/steam-props    │  │
│  │  - ... (more endpoints)          │  │
│  └──────────────────────────────────┘  │
└─────────────────────────────────────────┘
```

---

## Prerequisites

1. **Azure Account** with active subscription
2. **Azure CLI** installed: https://aka.ms/azure-cli
3. **GitHub Account** (for automatic deployments)
4. **Python 3.11** installed locally
5. **Azure Functions Core Tools** v4: https://aka.ms/azure-functions-core-tools

---

## Quick Deployment (Automated)

### Option 1: Deploy via Azure Portal (Recommended)

1. **Create Static Web App:**
   ```bash
   # Login to Azure
   az login

   # Set subscription (if you have multiple)
   az account set --subscription "YOUR_SUBSCRIPTION_NAME"

   # Create resource group
   az group create --name sigma-thermal-rg --location eastus2

   # Create Static Web App with GitHub integration
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

2. **Get deployment token:**
   ```bash
   az staticwebapp secrets list \
     --name sigma-thermal-calculators \
     --resource-group sigma-thermal-rg \
     --query "properties.apiKey" -o tsv
   ```

3. **Add GitHub Secret:**
   - Go to your GitHub repository
   - Settings → Secrets and variables → Actions
   - Click "New repository secret"
   - Name: `AZURE_STATIC_WEB_APPS_API_TOKEN`
   - Value: Paste the token from step 2
   - Click "Add secret"

4. **Push to GitHub:**
   ```bash
   git add .
   git commit -m "Add Azure deployment configuration"
   git push origin main
   ```

5. **Monitor deployment:**
   - Go to GitHub Actions tab
   - Watch the "Azure Static Web Apps CI/CD" workflow
   - Deployment typically takes 2-3 minutes

6. **Get your URL:**
   ```bash
   az staticwebapp show \
     --name sigma-thermal-calculators \
     --resource-group sigma-thermal-rg \
     --query "defaultHostname" -o tsv
   ```

   Your app will be at: `https://[generated-name].azurestaticapps.net`

---

### Option 2: Manual Deployment via VS Code

1. **Install Azure Static Web Apps Extension:**
   - Open VS Code
   - Install "Azure Static Web Apps" extension

2. **Sign in to Azure:**
   - Click Azure icon in sidebar
   - Sign in to your account

3. **Deploy:**
   - Right-click on `web` folder
   - Select "Deploy to Static Web App"
   - Follow prompts:
     - Create new or select existing
     - Name: `sigma-thermal-calculators`
     - Region: `East US 2`
     - App location: `/web`
     - API location: `/api`

---

## Project Structure

```
sigma-thermal/
├── web/                          # Static web content
│   ├── index.html
│   ├── resource.html
│   ├── css/
│   │   └── style.css
│   ├── js/
│   │   └── heating-value.js
│   ├── calculators/
│   │   └── heating-value.html
│   └── staticwebapp.config.json # Azure SWA configuration
├── api/                          # Azure Functions backend
│   ├── host.json                 # Functions host config
│   ├── requirements.txt          # Python dependencies
│   └── heating_value/            # Function: Heating Value Calculator
│       ├── __init__.py
│       └── function.json
└── .github/
    └── workflows/
        └── azure-static-web-apps.yml  # GitHub Actions deployment
```

---

## API Endpoints

### Heating Value Calculator

**Endpoint:** `POST /api/calculate/heating-value`

**Request:**
```json
{
  "ch4": 85.0,
  "c2h6": 10.0,
  "c3h8": 3.0,
  "c4h10": 1.0,
  "h2": 0.0,
  "co": 0.0,
  "h2s": 0.0,
  "co2": 1.0,
  "n2": 0.0
}
```

**Response:**
```json
{
  "hhv_mass": 22487.23,
  "lhv_mass": 20256.45,
  "hhv_volume": 1035.67,
  "lhv_volume": 932.89,
  "excel_comparison": {
    "hhv": 23875.0,
    "deviation": 0.0013
  }
}
```

---

## Configuration Files

### `web/staticwebapp.config.json`

Configures routing, headers, and MIME types for the static web app.

**Key settings:**
- Routes API requests to Azure Functions
- Sets up CORS headers
- Configures fallback routes
- Defines content security policy

### `api/host.json`

Configures Azure Functions runtime.

**Key settings:**
- Functions version: 2.0
- HTTP route prefix: `/api`
- Application Insights sampling
- Extension bundle version

### `.github/workflows/azure-static-web-apps.yml`

Automates deployment via GitHub Actions.

**Triggers:**
- Push to `main` branch
- Pull requests to `main`

**Actions:**
- Installs Python dependencies
- Builds sigma_thermal package
- Deploys static content to Azure
- Deploys API functions

---

## Custom Domain Setup (Optional)

### 1. Add Custom Domain

```bash
az staticwebapp hostname set \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --hostname calculators.yourdomain.com
```

### 2. Configure DNS

Add CNAME record to your DNS:
```
calculators.yourdomain.com  CNAME  [your-app].azurestaticapps.net
```

### 3. Verify and Enable SSL

SSL certificate is automatically provisioned by Azure (free).

---

## Environment Variables

### For Azure Functions

```bash
# Set environment variable for Functions
az staticwebapp appsettings set \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --setting-names \
    PYTHON_VERSION=3.11 \
    SCM_DO_BUILD_DURING_DEPLOYMENT=true
```

---

## Monitoring & Diagnostics

### View Logs

```bash
# Stream Function logs
az staticwebapp show \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg
```

### Application Insights

1. **Enable Application Insights:**
   ```bash
   az monitor app-insights component create \
     --app sigma-thermal-insights \
     --location eastus2 \
     --resource-group sigma-thermal-rg
   ```

2. **Link to Static Web App:**
   ```bash
   az staticwebapp appsettings set \
     --name sigma-thermal-calculators \
     --resource-group sigma-thermal-rg \
     --setting-names APPINSIGHTS_INSTRUMENTATIONKEY="YOUR_KEY"
   ```

---

## Cost Estimation

### Azure Static Web Apps Pricing (East US 2)

**Free Tier:**
- ✅ 100 GB bandwidth/month
- ✅ 2 custom domains
- ✅ Azure Functions included
- ✅ SSL certificates included
- ✅ GitHub integration
- **Cost:** $0/month

**Standard Tier** (if you need more):
- 100 GB bandwidth/month (additional $0.20/GB)
- Unlimited custom domains
- SLA: 99.95%
- **Cost:** $9/month + usage

**Recommended for this project:** Free Tier (sufficient for testing and moderate use)

---

## Testing Deployment

### 1. Test Static Pages

```bash
# Get your URL
URL=$(az staticwebapp show \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --query "defaultHostname" -o tsv)

# Test home page
curl https://$URL/

# Test resource page
curl https://$URL/resource.html
```

### 2. Test API Endpoint

```bash
# Test heating value calculation
curl -X POST https://$URL/api/calculate/heating-value \
  -H "Content-Type: application/json" \
  -d '{
    "ch4": 100.0, "c2h6": 0.0, "c3h8": 0.0, "c4h10": 0.0,
    "h2": 0.0, "co": 0.0, "h2s": 0.0, "co2": 0.0, "n2": 0.0
  }'
```

Expected response:
```json
{
  "hhv_mass": 23875.0,
  "lhv_mass": 21495.0,
  "hhv_volume": 1012.0,
  "lhv_volume": 910.0,
  "excel_comparison": {
    "hhv": 23875.0,
    "deviation": 0.0
  }
}
```

---

## Troubleshooting

### Issue: Function not deploying

**Solution:**
```bash
# Check if sigma_thermal package is installed
cd api
pip install -e ..

# Test function locally
func start
```

### Issue: CORS errors in browser

**Solution:** Check `staticwebapp.config.json` has correct headers:
```json
{
  "globalHeaders": {
    "content-security-policy": "default-src 'self' https: 'unsafe-inline' 'unsafe-eval'"
  }
}
```

### Issue: 404 on API endpoint

**Solution:** Verify route in `function.json`:
```json
{
  "route": "calculate/heating-value"
}
```

Full URL should be: `https://your-app.azurestaticapps.net/api/calculate/heating-value`

### Issue: Import errors in Azure Function

**Solution:** Add path setup in `__init__.py`:
```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
```

---

## Local Development

### Test Static Web App Locally

```bash
# Install Azure Static Web Apps CLI
npm install -g @azure/static-web-apps-cli

# Start local development server
swa start web --api-location api
```

Access at: `http://localhost:4280`

### Test Azure Functions Locally

```bash
cd api
func start
```

Access at: `http://localhost:7071/api/calculate/heating-value`

---

## Updating the Deployment

### Automatic (GitHub Actions)

Any push to `main` branch triggers automatic deployment:

```bash
git add .
git commit -m "Update calculator functionality"
git push origin main
```

### Manual

```bash
# Deploy using Azure CLI
az staticwebapp deploy \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --app-location ./web \
  --api-location ./api
```

---

## Security Considerations

1. **API Authentication:** Currently set to `anonymous`. Consider adding authentication:
   ```json
   {
     "authLevel": "function"
   }
   ```

2. **Rate Limiting:** Configure in `staticwebapp.config.json`

3. **Input Validation:** Already implemented in Azure Functions

4. **HTTPS:** Automatically enforced by Azure Static Web Apps

---

## Next Steps

1. ✅ Deploy to Azure Static Web Apps (East US 2)
2. ⏳ Add remaining calculator endpoints:
   - Air Requirement Calculator
   - Products of Combustion
   - Steam Properties
   - Water Properties
   - Flash Steam
3. ⏳ Configure custom domain (optional)
4. ⏳ Set up Application Insights monitoring
5. ⏳ Add authentication (if needed)

---

## Support & Documentation

- **Azure Static Web Apps Docs:** https://aka.ms/swa-docs
- **Azure Functions Python Docs:** https://aka.ms/azure-functions-python
- **GitHub Actions for Azure:** https://github.com/Azure/actions

---

## Summary

### What Was Created

✅ **Static Web App Configuration** (`web/staticwebapp.config.json`)
✅ **Azure Functions Backend** (`api/heating_value/`)
✅ **GitHub Actions Workflow** (`.github/workflows/azure-static-web-apps.yml`)
✅ **Deployment Documentation** (This file)

### Deployment Command

```bash
# One-line deployment (after Azure CLI login)
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

**Your app will be live at:** `https://[generated-name].azurestaticapps.net`

---

**Status:** Ready to deploy
**Region:** East US 2
**Estimated deployment time:** 3-5 minutes
**Cost:** Free tier (sufficient for development and testing)

