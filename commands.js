// ***********************************************
// This example commands.js shows you how to
// create various custom commands and overwrite
// existing commands.
//
// For more comprehensive examples of custom
// commands please read more here:
// https://on.cypress.io/custom-commands
// ***********************************************
//
//
// -- This is a parent command --
// Cypress.Commands.add('login', (email, password) => { ... })
//
//
// -- This is a child command --
// Cypress.Commands.add('drag', { prevSubject: 'element'}, (subject, options) => { ... })
//
//
// -- This is a dual command --
// Cypress.Commands.add('dismiss', { prevSubject: 'optional'}, (subject, options) => { ... })
//
//
// -- This will overwrite an existing command --
// Cypress.Commands.overwrite('visit', (originalFn, url, options) => { ... })

/// <reference types="Cypress"/>

// In cypress/support/commands.js


const xlsx = require('xlsx');

Cypress.Commands.add('readExcelFile', (filePath) => {
  return cy.readFile(filePath, 'binary').then((content) => {
    const workbook = xlsx.read(content, { type: 'binary' });
    // You can process the workbook and return relevant data
    // For example, you can return worksheets, rows, columns, etc.
    return workbook;
  });
});

Cypress.Commands.add('writeExcelFile', (filePath, data) => {
  const workbook = xlsx.utils.book_new();
  // Manipulate 'workbook' using xlsx library
  // Add worksheets, add data, etc.

  const binaryContent = xlsx.write(workbook, { type: 'binary' });
  cy.writeFile(filePath, binaryContent, 'binary');
});


// Below code is working fine for xlsx file reading.

import * as XLSX from 'xlsx';
 
Cypress.Commands.add('readExcelFile', (filePath) => {
  return cy.readFile(filePath, 'binary').then((content) => {
    const workbook = XLSX.read(content, { type: 'binary' });
    return workbook.Sheets[workbook.SheetNames[0]]; // Assuming you want the first sheet
  });
});

