# Blue Heron Dashboard — Production Deployment Guide

## Architecture
Browser → Azure Static Web App → /api/projects (Azure Function) → Fabric SQL Endpoint

## Step 1 — Push to GitHub

```bash
cd C:\BH_Dashboard\bh-dashboard-prod
git init
git add .
git commit -m "Initial Blue Heron dashboard"
git branch -M main
git remote add origin https://github.com/YOUR_ORG/bh-dashboard.git
git push -u origin main
```

## Step 2 — Create Azure Static Web App

1. Go to portal.azure.com
2. Search "Static Web Apps" → Create
3. Fill in:
   - Subscription: your subscription
   - Resource Group: rg-blueheron (create new)
   - Name: bh-project-dashboard
   - Plan: Free
   - Region: West US 2
   - Source: GitHub
   - Organization: your org
   - Repository: bh-dashboard
   - Branch: main
   - Build preset: React
   - App location: /
   - Api location: api
   - Output location: dist
4. Click Review + Create → Create
5. Azure will add the deploy workflow token to your GitHub repo automatically

## Step 3 — Add environment variables in Azure

In your Static Web App → Configuration → Application settings, add:

| Name | Value |
|------|-------|
| TENANT_ID | 7163abb0-3c0a-412b-b0c8-c9fdda5b1d3b |
| CLIENT_ID | 434f0a49-7cb8-4805-a62d-65a385169920 |
| CLIENT_SECRET | <your-client-secret> |
| FABRIC_ENDPOINT | wcvwg4ikhqvudmgizh65uwy5hm-ee2wcml4grlubetoiccdd5r7gq.datawarehouse.fabric.microsoft.com |
| DATABASE | lh_blueheron_dev |

## Step 4 — Enable Azure AD authentication (company-wide login)

In your Static Web App → Authentication:
1. Add identity provider → Microsoft
2. Use existing app registration:
   - Client ID: 434f0a49-7cb8-4805-a62d-65a385169920
   - Tenant ID: 7163abb0-3c0a-412b-b0c8-c9fdda5b1d3b
   - Client secret: <your-client-secret>
3. Restrict access: Require authentication

Now only users in your Azure AD tenant can access the dashboard.

## Step 5 — Done!

Your dashboard URL will be:
https://bh-project-dashboard.azurestaticapps.net

Every push to main branch auto-deploys in ~2 minutes.
Data auto-refreshes every 5 minutes from Fabric.

## Local development

```bash
npm install
npm run dev
# Runs on http://localhost:5173
# API calls proxy to localhost:7071 (needs Azure Functions Core Tools for local API)
```
