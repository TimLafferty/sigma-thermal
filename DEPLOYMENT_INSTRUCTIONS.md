# Azure Deployment Instructions

## Quick Deploy (5-10 minutes)

### Step 1: Create GitHub Repository

1. Go to https://github.com/new
2. Repository name: `sigma-thermal`
3. Description: "Professional thermal engineering calculation library with Excel UDFs, web calculators, and Azure Functions API"
4. Set to **Public** or **Private** (your choice)
5. DO NOT initialize with README (we already have one)
6. Click "Create repository"

### Step 2: Push Code to GitHub

From the terminal, run these commands (replace `YOUR_GITHUB_USERNAME`):

```bash
cd /Users/timlafferty/Repos/sigma-thermal

git remote add origin https://github.com/YOUR_GITHUB_USERNAME/sigma-thermal.git
git branch -M main
git push -u origin main
```

### Step 3: Create Azure Static Web App

**Option A: Using Azure CLI (Automated)**

Run the deployment script:
```bash
./deploy-azure.sh
```

When prompted:
- **Subscription**: Press Enter to use default "Azure subscription 1"
- **GitHub username**: Enter your GitHub username

The script will:
- Create resource group `sigma-thermal-rg` in East US 2
- Create Azure Static Web App `sigma-thermal-calculators`
- Set up GitHub Actions integration
- Provide deployment token

**Option B: Using Azure Portal (Manual)**

1. Go to https://portal.azure.com
2. Click "+ Create a resource"
3. Search for "Static Web App"
4. Click "Create"
5. Configuration:
   - **Subscription**: Azure subscription 1
   - **Resource Group**: Create new → `sigma-thermal-rg`
   - **Name**: `sigma-thermal-calculators`
   - **Region**: East US 2
   - **Deployment source**: GitHub
   - **GitHub account**: Click "Sign in with GitHub"
   - **Organization**: Your GitHub username
   - **Repository**: sigma-thermal
   - **Branch**: main
   - **Build Presets**: Custom
   - **App location**: `/web`
   - **Api location**: `/api`
   - **Output location**: (leave empty)
6. Click "Review + create"
7. Click "Create"

### Step 4: Get Deployment Token

After creation:
1. Go to resource `sigma-thermal-calculators`
2. Click "Manage deployment token" (or go to "Settings" → "Configuration")
3. Copy the deployment token

### Step 5: Add GitHub Secret

1. Go to https://github.com/YOUR_GITHUB_USERNAME/sigma-thermal/settings/secrets/actions
2. Click "New repository secret"
3. Name: `AZURE_STATIC_WEB_APPS_API_TOKEN`
4. Value: Paste the deployment token
5. Click "Add secret"

### Step 6: Trigger Deployment

Push any change to trigger deployment:
```bash
# Already done - the initial push will trigger deployment!
```

Or manually trigger:
1. Go to https://github.com/YOUR_GITHUB_USERNAME/sigma-thermal/actions
2. Click on "Azure Static Web Apps CI/CD" workflow
3. Click "Run workflow"

### Step 7: Monitor Deployment

1. Watch progress at: https://github.com/YOUR_GITHUB_USERNAME/sigma-thermal/actions
2. Deployment takes 2-3 minutes
3. Check for ✅ green checkmark

### Step 8: Access Your App

After successful deployment:
1. Go to https://portal.azure.com
2. Navigate to `sigma-thermal-calculators` resource
3. Find "URL" on overview page
4. Should look like: `https://[unique-name].azurestaticapps.net`

**Your app will include:**
- ✅ Web calculators at `/`
- ✅ Documentation site at `/docs/`
- ✅ API endpoints at `/api/`
- ✅ All resources and guides

---

## Troubleshooting

### GitHub Actions Workflow Fails

Check:
1. `AZURE_STATIC_WEB_APPS_API_TOKEN` secret is set correctly
2. GitHub has permission to deploy to Azure (check authorizations)
3. Build logs in Actions tab for specific errors

### 404 Errors on Deployed Site

Check:
1. `staticwebapp.config.json` is in `/web` directory ✅ (already there)
2. App location is set to `/web` in Azure configuration
3. Files deployed correctly (check deployment logs)

### API Functions Not Working

Check:
1. API location is set to `/api` in Azure configuration
2. `requirements.txt` exists in `/api` directory ✅ (already there)
3. Python version is 3.11 in workflow ✅ (already configured)

---

## What Gets Deployed

### Web Application (`/web` directory)
- **Home page**: `/index.html`
- **Calculators**: `/calculators/heating-value.html`, etc.
- **Resources**: `/resource.html`
- **Documentation**: `/docs/index.html` (NEW!)
  - Home overview
  - Technical resources
  - Project progress
  - Migration guide
  - Test results

### API Functions (`/api` directory)
- **Heating Value**: `/api/heating_value` (POST)
- Additional endpoints ready to add

### Static Assets
- CSS: `/css/style.css`, `/docs/css/docs.css`
- JavaScript: `/js/*.js`
- Configuration: `/staticwebapp.config.json`

---

## Post-Deployment

### Test Your Deployment

1. **Web Calculators**: Navigate to `https://[your-app].azurestaticapps.net`
2. **Documentation**: Navigate to `https://[your-app].azurestaticapps.net/docs/`
3. **API Health**: Test heating value calculator
4. **Navigation**: Click through all pages to verify links work

### Share Your App

Your app URL: `https://[unique-name].azurestaticapps.net`

Example documentation URLs:
- Documentation home: `/docs/index.html`
- Migration guide: `/docs/migration-guide.html`
- Test results: `/docs/test-results.html`

---

## Next Steps

1. ✅ Add custom domain (optional)
2. ✅ Set up staging environments (already configured in workflow)
3. ✅ Add more calculator pages
4. ✅ Expand API endpoints
5. ✅ Add authentication (future)

---

**Quick Deploy Checklist:**

- [ ] Create GitHub repository
- [ ] Push code to GitHub
- [ ] Create Azure Static Web App (via script or portal)
- [ ] Add GitHub secret for deployment token
- [ ] Watch GitHub Actions deployment
- [ ] Access deployed app
- [ ] Test all pages and calculators
- [ ] Share app URL with team

**Estimated Time:** 5-10 minutes
**Cost:** Free tier (100 GB bandwidth/month, no credit card needed for free tier)
