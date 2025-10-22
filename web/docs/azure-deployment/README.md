# Azure Deployment Documentation

**Deploy Sigma Thermal web calculators to Azure Static Web Apps**

---

## Overview

Deploy professional web-based thermal calculators to Azure with:

- ✅ **Static web hosting** for HTML/CSS/JS
- ✅ **Azure Functions** backend API
- ✅ **Automated CI/CD** via GitHub Actions
- ✅ **Free tier available** (100 GB bandwidth/month)
- ✅ **SSL certificates** included
- ✅ **Custom domains** supported
- ✅ **East US 2** region

---

## Quick Start

### 5-Minute Deployment

```bash
./deploy-azure.sh
```

This script will:
1. Login to Azure
2. Create resource group in East US 2
3. Create Static Web App
4. Connect to GitHub
5. Provide deployment token

**Full guide:** [Quick Start](quick-start.md)

---

## Documentation Files

### [Quick Start](quick-start.md)

**5-minute deployment guide**

Topics covered:
- Prerequisites
- One-command deployment
- Manual deployment steps (alternative)
- Testing deployment
- Cost information
- Troubleshooting

**Start here for fastest deployment.**

### [Deployment Guide](deployment-guide.md)

**Complete Azure setup and configuration**

Topics covered:
- Architecture overview
- Detailed deployment steps
- API endpoint documentation
- Configuration files explained
- Custom domain setup
- Environment variables
- Monitoring with Application Insights
- Cost estimation
- Testing procedures
- Security considerations
- Troubleshooting

**Reference guide for advanced configuration.**

---

## Architecture

```
┌─────────────────────────────────────────┐
│   Azure Static Web Apps (East US 2)    │
│                                         │
│  ┌──────────────────────────────────┐  │
│  │  Static Content (HTML/CSS/JS)    │  │
│  │  - index.html                    │  │
│  │  - resource.html                 │  │
│  │  - calculators/*.html            │  │
│  └──────────────────────────────────┘  │
│                                         │
│  ┌──────────────────────────────────┐  │
│  │  Azure Functions API             │  │
│  │  - /api/calculate/heating-value  │  │
│  │  - /api/calculate/steam-props    │  │
│  └──────────────────────────────────┘  │
└─────────────────────────────────────────┘
```

---

## What Gets Deployed

### Static Web Content

Located in `web/` directory:
- `index.html` - Landing page
- `resource.html` - Technical reference page
- `calculators/*.html` - Calculator forms
- `css/style.css` - Professional styling
- `js/*.js` - Form handling and API calls

### Azure Functions API

Located in `api/` directory:
- `heating_value/` - Heating value calculator endpoint
- `host.json` - Functions runtime configuration
- `requirements.txt` - Python dependencies

### CI/CD Pipeline

Located in `.github/workflows/`:
- `azure-static-web-apps.yml` - Automatic deployment on push

### Configuration

- `web/staticwebapp.config.json` - Azure SWA routing and CORS
- `deploy-azure.sh` - Deployment automation script

---

## Prerequisites

1. **Azure Account** - https://azure.microsoft.com/free/
2. **Azure CLI** - https://aka.ms/azure-cli
3. **GitHub Account** - For CI/CD integration
4. **Python 3.11+** - For local testing

---

## Deployment Steps (Summary)

### 1. Install Azure CLI

```bash
# macOS
brew install azure-cli

# Windows
winget install Microsoft.AzureCLI

# Linux
curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash
```

### 2. Run Deployment Script

```bash
./deploy-azure.sh
```

### 3. Configure GitHub Secret

Add the deployment token to GitHub:
1. Go to: `https://github.com/YOUR_USERNAME/sigma-thermal/settings/secrets/actions`
2. Create secret: `AZURE_STATIC_WEB_APPS_API_TOKEN`
3. Paste token from deployment script

### 4. Deploy Code

```bash
git add .
git commit -m "Deploy to Azure"
git push origin main
```

### 5. Access Your App

After 2-3 minutes:
```
https://[generated-name].azurestaticapps.net
```

**Full details:** [Quick Start](quick-start.md)

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
  "hhv_mass": 23389.33,
  "lhv_mass": 20256.45,
  "hhv_volume": 1009.1,
  "lhv_volume": 907.8
}
```

**Full API docs:** [Deployment Guide](deployment-guide.md#api-endpoints)

---

## Testing Deployment

### Test Static Page

```bash
curl https://YOUR-APP-URL.azurestaticapps.net/
```

### Test API Endpoint

```bash
curl -X POST https://YOUR-APP-URL.azurestaticapps.net/api/calculate/heating-value \
  -H "Content-Type: application/json" \
  -d '{"ch4":100,"c2h6":0,"c3h8":0,"c4h10":0,"h2":0,"co":0,"h2s":0,"co2":0,"n2":0}'
```

Expected: JSON response with heating values

---

## Cost

### Free Tier (Recommended)

- ✅ 100 GB bandwidth/month
- ✅ SSL certificates included
- ✅ Custom domains included
- ✅ Azure Functions included
- **Cost: $0/month**

### Standard Tier

- 100 GB bandwidth/month (additional $0.20/GB)
- Unlimited custom domains
- SLA: 99.95%
- **Cost: $9/month + usage**

**Recommendation:** Free tier is sufficient for testing and moderate use

---

## Custom Domain Setup

### 1. Add Custom Domain

```bash
az staticwebapp hostname set \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --hostname calculators.yourdomain.com
```

### 2. Configure DNS

Add CNAME record:
```
calculators.yourdomain.com  CNAME  [your-app].azurestaticapps.net
```

### 3. SSL Certificate

Automatically provisioned by Azure (free)

**Full guide:** [Deployment Guide](deployment-guide.md#custom-domain-setup)

---

## Monitoring

### View Deployment Status

GitHub Actions tab:
```
https://github.com/YOUR_USERNAME/sigma-thermal/actions
```

### View Application Logs

Azure Portal → Static Web App → Logs

### Set Up Application Insights

```bash
az monitor app-insights component create \
  --app sigma-thermal-insights \
  --location eastus2 \
  --resource-group sigma-thermal-rg
```

**Full guide:** [Deployment Guide](deployment-guide.md#monitoring--diagnostics)

---

## Troubleshooting

### Issue: "az: command not found"

**Solution:** Install Azure CLI: https://aka.ms/azure-cli

### Issue: GitHub authentication fails

**Solution:** Run `az login` again and authorize GitHub access

### Issue: API returns 500 error

**Solution:**
1. Check Application Insights logs
2. Test locally: `cd api && func start`
3. Verify Python dependencies in `requirements.txt`

### Issue: 404 on API endpoint

**Solution:** Verify route in `function.json`:
```json
{
  "route": "calculate/heating-value"
}
```

Full URL: `https://your-app.azurestaticapps.net/api/calculate/heating-value`

**Full troubleshooting:** [Deployment Guide](deployment-guide.md#troubleshooting)

---

## Local Development

### Test Static Web App Locally

```bash
# Install Azure Static Web Apps CLI
npm install -g @azure/static-web-apps-cli

# Start local dev server
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

## Updating Deployment

### Automatic (GitHub Actions)

Any push to `main` branch triggers automatic deployment:

```bash
git add .
git commit -m "Update calculator functionality"
git push origin main
```

### Manual

```bash
az staticwebapp deploy \
  --name sigma-thermal-calculators \
  --resource-group sigma-thermal-rg \
  --app-location ./web \
  --api-location ./api
```

---

## Next Steps

### After Deployment

1. ✅ Test all calculator endpoints
2. ⏳ Add remaining API endpoints (air requirement, steam properties, etc.)
3. ⏳ Configure custom domain (optional)
4. ⏳ Set up Application Insights monitoring
5. ⏳ Add authentication (if needed)

---

## Files and Configuration

### Key Files

| File | Description |
|------|-------------|
| `web/staticwebapp.config.json` | Azure SWA configuration |
| `api/host.json` | Functions runtime config |
| `api/heating_value/__init__.py` | Heating value API endpoint |
| `.github/workflows/azure-static-web-apps.yml` | CI/CD pipeline |
| `deploy-azure.sh` | Deployment automation script |

**Full details:** [Deployment Guide](deployment-guide.md#project-structure)

---

## Support

### Resources

- **Azure Static Web Apps Docs:** https://aka.ms/swa-docs
- **Azure Functions Python Docs:** https://aka.ms/azure-functions-python
- **GitHub Actions for Azure:** https://github.com/Azure/actions

### Getting Help

1. Check [Quick Start](quick-start.md) troubleshooting
2. Review [Deployment Guide](deployment-guide.md) for detailed docs
3. Check Azure Portal logs
4. Contact GTS Energy Inc.

---

## Summary

### What You Get

- ✅ Professional web calculators
- ✅ Azure Static Web Apps hosting
- ✅ Azure Functions API backend
- ✅ Automated CI/CD pipeline
- ✅ Free tier available
- ✅ SSL certificates included
- ✅ Custom domain support

### Deployment Time

- **Setup:** 5 minutes
- **First deployment:** 2-3 minutes
- **Updates:** Automatic on push

### Cost

- **Free tier:** $0/month (recommended for testing)
- **Standard tier:** $9/month + usage (for production)

---

**Back to:** [Main Documentation](../README.md)
