#!/usr/bin/env node

// File: reorganize.js
// Purpose: Reorganizes the repository structure by creating directories and moving files

const fs = require('fs');
const path = require('path');

// Structure definition - maps current files to their new locations
const fileMapping = {
  'config-module.js': 'src/config/config.gs',
  'menu-module.js': 'src/core/menu.gs',
  'protection-module.js': 'src/core/protection.gs',
  'utilities-module.js': 'src/core/utilities.gs',
  'instructions-module.js': 'src/sheets/basic/instructions.gs',
  'datatables-module.js': 'src/sheets/basic/dataTables.gs',
  'trip-settings-module.js': 'src/sheets/basic/tripSettings.gs',
  'lodging-module.js': 'src/sheets/basic/lodging.gs',
  'transportation-module.js': 'src/sheets/advanced/transportation.gs',
  'activities-module.js': 'src/sheets/advanced/activities.gs',
  'dashboard-module.js': 'src/sheets/advanced/dashboard.gs',
  'main-module.js': 'src/main.gs'
};

// Create required directories
function createDirectories() {
  const directories = [
    'src/config',
    'src/core',
    'src/sheets/basic',
    'src/sheets/advanced'
  ];

  directories.forEach(dir => {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
      console.log(`Created directory: ${dir}`);
    }
  });
}

// Move and rename files
function moveFiles() {
  Object.entries(fileMapping).forEach(([oldPath, newPath]) => {
    try {
      if (fs.existsSync(oldPath)) {
        fs.renameSync(oldPath, newPath);
        console.log(`Moved ${oldPath} to ${newPath}`);
      } else {
        console.warn(`Warning: Source file ${oldPath} not found`);
      }
    } catch (error) {
      console.error(`Error moving ${oldPath}: ${error.message}`);
    }
  });
}

// Main execution
function main() {
  console.log('Starting repository reorganization...');
  
  // Create backup of current structure
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backupDir = `backup-${timestamp}`;
  fs.mkdirSync(backupDir);
  
  // Backup existing files
  Object.keys(fileMapping).forEach(file => {
    if (fs.existsSync(file)) {
      fs.copyFileSync(file, path.join(backupDir, file));
    }
  });
  
  console.log(`Created backup in directory: ${backupDir}`);
  
  // Create new directory structure
  createDirectories();
  
  // Move and rename files
  moveFiles();
  
  console.log('Repository reorganization complete!');
}

// Execute the script
main();
