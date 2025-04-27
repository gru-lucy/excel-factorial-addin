# Excel Factorial Add-in

An Excel custom functions add-in built with TypeScript, React, and the Office JavaScript API. This sample demonstrates how to create synchronous, streaming, and spill functions for Excel, including a factorial spill function with interactive orientation control.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
  - [Install Dependencies](#install-dependencies)
  - [Trust Developer Certificate](#trust-developer-certificate)
  - [Run Locally](#run-locally)
  - [Sideload to Excel](#sideload-to-excel)
- [Custom Functions](#custom-functions)
- [Task Pane](#task-pane)
- [Scripts](#scripts)
- [Building for Production](#building-for-production)
- [Debugging](#debugging)
- [Linting & Formatting](#linting--formatting)
- [Contributing](#contributing)
- [Support](#support)
- [License](#license)
- [Repository](#repository)

## Features

- **ADD:** Add two numbers.
- **CLOCK:** Display the current time, updating every second (streaming).
- **INCREMENT:** Increment a counter by a specified value every second (streaming).
- **LOG:** Write a message to the console and return it.
- **FACTORIALROW:** Spill factorials from 1! up to N! in a row or column, with orientation controlled via the task pane.

## Prerequisites

- [Node.js](https://nodejs.org/) v14 or later
- [npm](https://www.npmjs.com/) v6 or later
- Microsoft Excel (Office 365 or Excel 2016+ with Office.js support)
- Optional: [Office Add-in CLI](https://github.com/OfficeDev/Office-Addin-CLI)

## Getting Started

### Install Dependencies

```bash
npm install
```

### Trust Developer Certificate

```bash
# (Optional) Install and trust the HTTPS certificate
npx office-addin-dev-certs install --trust
```

### Run Locally

```bash
npm run dev-server
```

The add-in will be served at `https://localhost:3000`.

### Sideload to Excel

1. Open Excel.
2. Go to **Insert > Office Add-ins > Manage My Add-ins**.
3. Click **Upload My Add-in** and select `manifest.xml`.

## Custom Functions

Use the `TESTVELIXO` namespace for all custom functions:

```excel
=TESTVELIXO.ADD(1, 2)
=TESTVELIXO.CLOCK()
=TESTVELIXO.INCREMENT(5)
=TESTVELIXO.LOG("Hello World")
=TESTVELIXO.FACTORIALROW(7)
```

## Task Pane

Open the task pane in Excel by clicking the **Show Task Pane** button in the **Home** tab. Use the radio buttons to switch the orientation of the `FACTORIALROW` output between a row and a column. Changes are saved in `localStorage` and trigger an automatic recalculation.

## Scripts

- `npm run build` - Build the add-in for production.
- `npm run build:dev` - Build the add-in in development mode.
- `npm run dev-server` - Start the webpack dev server.
- `npm run watch` - Watch for file changes and rebuild.
- `npm run start` - Sideload and run the add-in in Excel.
- `npm run stop` - Stop the debugging session.
- `npm run lint` - Run lint checks.
- `npm run lint:fix` - Automatically fix lint issues.
- `npm run prettier` - Format code using Prettier.
- `npm run validate` - Validate the `manifest.xml` file.
- `npm run signin` - Sign in to a Microsoft 365 account.
- `npm run signout` - Sign out of the Microsoft 365 account.

## Building for Production

```bash
npm run build
```

## Debugging

```bash
npm run start
```

## Linting & Formatting

```bash
npm run lint
npm run lint:fix
npm run prettier
```

## Contributing

Please read [CODE_OF_CONDUCT.md](./CODE_OF_CONDUCT.md) before contributing.

## Support

See [SUPPORT.md](./SUPPORT.md) for support and troubleshooting.

## License

This project is licensed under the MIT License.

## Repository

https://github.com/OfficeDev/Excel-Custom-Functions
