# GitHub Setup Guide

This guide will help you upload your PowerPoint Automation Agent project to GitHub.

## Step 1: Create a GitHub Repository

1. **Go to GitHub.com** and sign in to your account
2. **Click the "+" icon** in the top right corner
3. **Select "New repository"**
4. **Fill in the repository details:**
   - **Repository name:** `powerpoint-automation-agent` (or your preferred name)
   - **Description:** `AI-powered PowerPoint automation agent with MCP server for mathematical problem solving and visualization`
   - **Visibility:** Choose Public or Private
   - **DO NOT** initialize with README, .gitignore, or license (we already have these)
5. **Click "Create repository"**

## Step 2: Connect Local Repository to GitHub

After creating the repository, GitHub will show you commands. Use these commands in your terminal:

```bash
# Add the remote repository (replace YOUR_USERNAME with your GitHub username)
git remote add origin https://github.com/YOUR_USERNAME/powerpoint-automation-agent.git

# Set the main branch name
git branch -M main

# Push your code to GitHub
git push -u origin main
```

## Step 3: Verify Upload

1. **Refresh your GitHub repository page**
2. **You should see all your files:**
   - `powerpoint_working_agent.py`
   - `powerpoint_working_mcp_server.py`
   - `README.md`
   - `requirements.txt`
   - `.gitignore`
   - `LICENSE`
   - `CONTRIBUTING.md`

## Step 4: Add Repository Topics (Optional)

1. **Go to your repository page**
2. **Click the gear icon** next to "About"
3. **Add topics** like:
   - `python`
   - `automation`
   - `powerpoint`
   - `mcp-server`
   - `ai-agent`
   - `mathematical-computing`

## Step 5: Create a Release (Optional)

1. **Go to "Releases"** in your repository
2. **Click "Create a new release"**
3. **Tag version:** `v1.0.0`
4. **Release title:** `PowerPoint Automation Agent v1.0.0`
5. **Description:** Copy from your README.md
6. **Click "Publish release"**

## Repository Structure

Your GitHub repository will have this clean structure:

```
powerpoint-automation-agent/
├── .gitignore
├── CONTRIBUTING.md
├── LICENSE
├── README.md
├── requirements.txt
├── powerpoint_working_agent.py
└── powerpoint_working_mcp_server.py
```

## Next Steps

1. **Share your repository** with your trainer
2. **Clone the repository** on other machines:
   ```bash
   git clone https://github.com/YOUR_USERNAME/powerpoint-automation-agent.git
   ```
3. **Continue development** by making changes and pushing updates

## Troubleshooting

### If you get authentication errors:
- Use GitHub CLI: `gh auth login`
- Or use Personal Access Token instead of password

### If you get permission errors:
- Make sure you're the owner of the repository
- Check that the repository name matches exactly

### If files don't appear:
- Check that you ran `git add .` and `git commit` before pushing
- Verify the remote URL is correct

## Success! 🎉

Your PowerPoint Automation Agent is now on GitHub and ready to share with your trainer!
