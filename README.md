# Excel Add-in Vite Template

A modern, efficient template for building Excel add-ins using Vite, React, and TypeScript.

![Excel Add-in with Vite](https://img.shields.io/badge/Excel%20Add--in-Vite-blue)
![React](https://img.shields.io/badge/React-18-blue)
![TypeScript](https://img.shields.io/badge/TypeScript-5-blue)
![License](https://img.shields.io/badge/License-MIT-green)

## 🚀 Features

- **Vite** - Lightning fast HMR and build times
- **React** - Component-based UI development
- **TypeScript** - Type safety and better developer experience
- **Fluent UI React** - Microsoft's design system
- **Office JS API Integration** - Pre-configured for Excel
- **Dev Certificates** - Automated HTTPS setup
- **Hot Module Replacement** - See changes instantly

## 📋 Prerequisites

- [Node.js](https://nodejs.org) (LTS version recommended)
- [Yarn](https://yarnpkg.com/) package manager
- Microsoft Excel (desktop or online)

## 🛠️ Getting Started

### Creating a New Project

#### Option 1: Use GitHub Template

1. Click the "Use this template" button at the top of this repository
2. Name your new repository
3. Clone your new repository locally

#### Option 2: Use GitHub CLI

```bash
gh repo create my-excel-addin --template username/excel-addin-vite-template
cd my-excel-addin
```

### Setting Up Your Project

1. Install dependencies:

   ```bash
   yarn install
   ```

2. Generate a development certificate (one-time setup):

   ```bash
   yarn generate-cert
   ```

3. Update the GUID in the manifest file:

   - Open `manifest.xml`
   - Replace `YOUR-GUID-HERE` with a unique GUID
   - You can generate a GUID with:
     ```bash
     node -e "console.log(require('crypto').randomUUID())"
     ```

4. Add your logo/icon files to the `assets` folder:
   - icon-16.png
   - icon-32.png
   - icon-80.png

## 💻 Development

Start the development server:

```bash
yarn start
```

This launches a dev server at `https://localhost:3000` with hot module replacement.

## 📥 Sideloading the Add-in

### Windows

Run the sideloading command:

```bash
yarn sideload
```

### macOS

1. Place your manifest file in:

   ```
   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```

   (You may need to create this directory)

2. Restart Excel

### Excel on the Web

1. Open Excel on the web
2. Go to Insert > Add-ins > Upload My Add-in
3. Browse to your manifest.xml file

## 🏗️ Building for Production

Create a production build:

```bash
yarn build
```

The built files will be in the `dist` directory.

## 🔍 Troubleshooting

### Common Issues

- **"This add-in is no longer available"**: Check that your development server is running and URLs in manifest are correct
- **Missing add-in in Excel**: Verify that sideloading was successful and manifest is valid
- **Certificate errors**: Run `yarn generate-cert` and ensure certificates are trusted
- **Blank taskpane**: Check console for errors and verify paths in the manifest

### Validation

Validate your manifest:

```bash
yarn validate
```

## 📁 Project Structure

```
excel-addin-vite/
├── assets/               # Static assets like icons
├── src/
│   ├── taskpane/        # UI components and code
│   │   ├── taskpane.html
│   │   ├── main.tsx
│   │   ├── Taskpane.tsx
│   │   └── taskpane.css
│   ├── functions/       # Office JS functions
│   │   ├── function-file.html
│   │   └── functions.ts
│   └── env.d.ts         # TypeScript declaration file
├── manifest.xml         # Add-in definition
├── tsconfig.json        # TypeScript configuration
├── vite.config.ts       # Vite configuration
└── package.json         # Project dependencies
```

## 📚 Additional Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Excel JavaScript API Reference](https://docs.microsoft.com/javascript/api/excel)
- [Vite Documentation](https://vitejs.dev/guide/)
- [React Documentation](https://react.dev/)
- [Fluent UI React](https://developer.microsoft.com/en-us/fluentui)

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Open a Pull Request
