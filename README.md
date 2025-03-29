# Excel Add-in with Vite

This is a modern Excel add-in project using Vite, React, and TypeScript.

## Features

- **Vite** - Fast, modern frontend build tool
- **React** - UI library
- **TypeScript** - Type safety
- **Fluent UI React** - Microsoft's design language
- **Hot Module Replacement** - Fast development experience

## Prerequisites

- [Node.js](https://nodejs.org) (LTS version recommended)
- [Yarn](https://yarnpkg.com/) package manager
- Microsoft Excel (desktop or online)

## Project Setup

1. Clone or download this project to your local machine

2. Install dependencies:

   ```bash
   cd excel-addin-vite
   yarn
   ```

3. Generate a development certificate (one-time setup):

   ```bash
   yarn generate-cert
   ```

4. Update the GUID in the manifest file:

   - Open `manifest.xml`
   - Replace `YOUR-GUID-HERE` with a unique GUID
   - You can generate a GUID with:
     ```bash
     node -e "console.log(require('crypto').randomUUID())"
     ```

5. Add your logo/icon files to the `assets` folder:
   - icon-16.png
   - icon-32.png
   - icon-80.png

## Development

Start the development server:

```bash
yarn start
```

This will launch a development server at `https://localhost:3000` with hot module replacement enabled.

## Sideloading the Add-in

### For Excel Desktop

1. Run the following command to start Excel with your add-in:

   ```bash
   yarn sideload
   ```

2. If the above command doesn't work, manually sideload:
   - Open Excel
   - Go to **Insert** > **Add-ins** > **My Add-ins**
   - Click **Upload My Add-in** and browse to your manifest.xml file

### For Excel Online

1. Open Excel Online
2. Go to **Insert** > **Office Add-ins** > **Upload My Add-in**
3. Browse to your manifest.xml file

## Building for Production

Create a production build:

```bash
yarn build
```

The built files will be in the `dist` directory.

## Troubleshooting

If you encounter issues with the manifest file:

1. Ensure your development server is running
2. Validate your manifest with `yarn validate`
3. Check that the URLs in the manifest match your Vite server configuration
4. Make sure your SSL certificate is properly installed
5. Look at browser console errors

## Project Structure

- `src/taskpane/` - UI components and code
- `src/functions/` - Office JavaScript functions
- `assets/` - Static assets like icons
- `manifest.xml` - Add-in definition
- `vite.config.ts` - Vite configuration

## Why Vite?

Vite offers several advantages over webpack for Excel add-in development:

1. **Faster startup and build times** - Vite uses native ES modules during development
2. **Simpler configuration** - Less boilerplate code to maintain
3. **Hot Module Replacement (HMR)** - See changes instantly without full page reloads
4. **Optimized builds** - Efficient production bundles with automatic code-splitting
5. **Built-in TypeScript support** - No need for additional loaders

## Additional Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Excel JavaScript API Reference](https://docs.microsoft.com/javascript/api/excel)
- [Vite Documentation](https://vitejs.dev/guide/)
- [React Documentation](https://reactjs.org/)
- [Fluent UI React](https://developer.microsoft.com/en-us/fluentui)
