/**
 * combine-modules.js
 * Combines all Google Apps Script modules into a single Code.gs file
 * in the correct dependency order.
 */

const fs = require('fs');
const path = require('path');

// Module order based on dependencies
const MODULE_ORDER = [
  // Core modules
  'src/config/config.gs',              // Global configuration
  'src/core/utilities.gs',             // Utility functions
  'src/core/protection.gs',            // Protection utilities
  'src/core/menu.gs',                  // Menu management
  
  // Basic sheet modules
  'src/sheets/basic/instructions.gs',   // Instructions sheet
  'src/sheets/basic/dataTables.gs',    // Reference data
  'src/sheets/basic/tripSettings.gs',   // Trip settings
  'src/sheets/basic/lodging.gs',       // Lodging comparison
  
  // Advanced sheet modules
  'src/sheets/advanced/transportation.gs', // Transportation
  'src/sheets/advanced/activities.gs',     // Activities
  'src/sheets/advanced/dashboard.gs',      // Dashboard
  
  // Main module (must be last)
  'src/main.gs'                        // Main entry point
];

/**
 * Combines all modules into a single Code.gs file
 * @param {string} outputPath - Path to output the combined file
 */
function combineModules(outputPath) {
  let combinedContent = '';
  
  // Add file header
  combinedContent += '/**\n';
  combinedContent += ' * Code.gs\n';
  combinedContent += ' * Combined Google Apps Script modules for Travel Budget Planner\n';
  combinedContent += ' * Generated on: ' + new Date().toISOString() + '\n';
  combinedContent += ' */\n\n';
  
  // Process each module in order
  MODULE_ORDER.forEach((modulePath) => {
    console.log(`Processing: ${modulePath}`);
    try {
      const content = fs.readFileSync(modulePath, 'utf8');
      
      // Add module separator
      combinedContent += '\n// ' + '='.repeat(80) + '\n';
      combinedContent += `// Module: ${path.basename(modulePath)}\n`;
      combinedContent += '// ' + '='.repeat(80) + '\n\n';
      
      // Add module content
      combinedContent += content + '\n';
      
    } catch (error) {
      console.error(`Error processing ${modulePath}:`, error);
      process.exit(1);
    }
  });
  
  // Write combined file
  try {
    fs.writeFileSync(outputPath, combinedContent);
    console.log(`Successfully created: ${outputPath}`);
  } catch (error) {
    console.error('Error writing output file:', error);
    process.exit(1);
  }
}

// Execute if run directly
if (require.main === module) {
  const outputPath = path.join(__dirname, 'Code.gs');
  combineModules(outputPath);
}

module.exports = combineModules;
