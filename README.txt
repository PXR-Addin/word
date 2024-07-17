# Office Add-in Development and Deployment Guide

This guide explains how to manage development and production environments for your Office Add-in.

## Project Structure

- `manifest.dev.xml`: Development manifest with localhost URLs
- `manifest.prod.xml`: Production manifest with SharePoint URLs
- `webpack.config.js`: Configured to handle both dev and prod builds
- `package.json`: Contains scripts for different build and start commands

## Development Workflow

1. Use `manifest.dev.xml` for local development.
2. Run `npm run start` to test the add-in locally.
3. Use `npm run build:dev` for development builds.
4. Use `npm run build:prod` for production builds.
5. Use `npm install -g office-addin-manifest manifest.xml` to check the manifest file

## Production Workflow

1. Ensure `manifest.prod.xml` has correct SharePoint URLs.
2. Run `npm run build:prod` for production builds.
3. Upload the `manifest.xml` from the `dist` folder to Microsoft 365 Admin Center.

## Key Points

- Always keep `manifest.dev.xml` and `manifest.prod.xml` in sync, except for URLs.
- Use the manifest from the `dist` folder for deployment, not the source files.
- The `webpack.config.js` is set up to copy the correct manifest based on the build type.

## npm Scripts

- `npm run start`: Starts the add-in for local development
- `npm run build:dev`: Builds the project for development
- `npm run build:prod`: Builds the project for production
- `npm run start:prod`: Starts the add-in with production settings (for testing)

## Deployment Steps

1. Run `npm run build:prod`
2. Navigate to Microsoft 365 Admin Center
3. Go to Settings > Integrated apps > Upload custom apps
4. Upload the `manifest.xml` from the `dist` folder

## Best Practices

- Always rebuild (`npm run build:prod`) before uploading a new manifest for deployment.
- Regularly compare dev and prod manifests to ensure they haven't unintentionally diverged.
- Keep both manifest files in version control, but ignore the `dist` folder.
- Test thoroughly in both development and production environments before final deployment.

## Troubleshooting

If you encounter issues with manifest generation or URL mismatches:
1. Check `webpack.config.js` for correct manifest handling
2. Verify that `manifest.dev.xml` and `manifest.prod.xml` are correctly set up
3. Ensure all URLs in the prod manifest point to your SharePoint locations

For any other issues, refer to the Office Add-in documentation or seek support from the development team.


--- Starting the Dev Environment ---

Inputs: Visual Code; npm

1. Open cmd, navigate to
cd "KaiWordTools"

2. Start the Dev Local Version using

npm run start

--- Building ---
Builds are output to the */dist Folder. There are two manifests: prod and dev. The only difference is that dev has localhost URLs. Use these commands to build the correct version in */dist - it should only change the manifest.xml file:

npm run build:prod
npm run build:dev

Remember:

Always use the manifest from the dist folder after a production build for deployment.
Never upload manifest.dev.xml or the source manifest.prod.xml to the Admin Center.
If you make changes to your add-in, always rebuild (npm run build:prod) before uploading a new manifest.