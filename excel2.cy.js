describe('Excel File Reading', () => {
    it('should read data from Excel file', () => {
      const filePath = 'cypress/fixtures/Template_PA Result.xlsm'; // Path relative to the project root
      
      cy.readExcelFile(filePath).then((sheet) => {
        // Define an array of cell references to validate
        const cellsToValidate = ['A1', 'B1'];
   
        // Iterate over the cells and perform validations
        for (const cellRef of cellsToValidate) {
          const expectedValue = getExpectedValue(cellRef);
   
          // Example: Asserting the value in the current cell
          expect(sheet[cellRef].v).to.equal(expectedValue);
        }
      });
    });
});

  // Helper function to get expected values based on cell reference
  function getExpectedValue(cellRef) {
    // Implement your logic to get the expected value based on the cell reference
    // For simplicity, return a static value in this example
    switch (cellRef) {
      case 'A1':
        return 'Header';
      case 'B1':
        return 'Reference';
   
      default:
        return 'DefaultValue';
    }
  }