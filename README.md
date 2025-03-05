# Excel Data Cleaner Add-in

A user-friendly Excel add-in for data cleaning with no technical knowledge required.

## Features

### Basic Features (Trial Version)
- **Remove Duplicates** - Find and remove duplicate rows
- **Trim Spaces** - Remove extra spaces from text cells
- **Text Case Conversion** - Convert text to UPPERCASE, lowercase, or Proper Case
- **Find & Replace** - Search and replace text across multiple cells

### Premium Features
- **Date Standardization** - Convert dates to a consistent format (MM/DD/YYYY, DD/MM/YYYY, or YYYY-MM-DD)
- **One-Click Cleanup** - Apply multiple cleaning operations at once
- **Number Formatting** - Fix numbers stored as text
- **Unlimited Usage** - Remove trial version limitations

## Trial Limitations
- Limited to 10 operations
- Duplicate removal limited to 100 rows
- Text case conversion limited to 2 columns

## Development

### Prerequisites
- Node.js (version 14 or later)
- npm (version 6 or later)
- Visual Studio Code (recommended)

### Getting Started
1. Clone the repository
2. Install dependencies:
   ```
   npm install
   ```
3. Start the development server:
   ```
   npm start
   ```
4. Excel will open with the add-in sideloaded

### Project Structure
```
excel-data-cleaner/
│
├── src/                     # Source files
│   ├── taskpane/           # Taskpane-related files
│   │   ├── taskpane.html   # Main HTML template
│   │   ├── taskpane.css    # Main CSS stylesheet
│   │   └── taskpane.js     # Main entry point
│   │
│   ├── features/           # Feature-specific code
│   │   ├── basic/          # Basic (trial) features
│   │   └── premium/        # Premium features
│   │
│   ├── utils/              # Utility functions
│   │
│   └── index.js            # Webpack entry point
│
├── assets/                 # Static assets (icons, etc.)
│
├── manifest.xml            # Add-in manifest
├── package.json            # NPM package configuration
├── webpack.config.js       # Webpack configuration
└── README.md               # Project documentation
```

### Building for Production
To build the add-in for production use:
```
npm run build
```

This will create a production build in the `dist` folder.

## Deployment

### Office Add-in Store
1. Create a production build
2. Submit your add-in to the [Microsoft Partner Center](https://partner.microsoft.com/en-us/dashboard/office/overview)

### Sideloading
Follow the instructions in [Sideload Office Add-ins for testing](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).

## License
This project is licensed under the MIT License - see the LICENSE file for details.