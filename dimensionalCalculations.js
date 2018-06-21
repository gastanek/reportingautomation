function calcArrAmounts() {
  
    //aggregates dimensional values 
  
    var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Win Data');
    var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Solution-Use Case Matrix');
  
    var solutionsArray = new Array('S', 'R', 'T', 'U', 'V', 'W', 'X')
    var useCaseArray = new Array('Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN');
    var targetSolutionsArray = new Array('15', '16', '17', '18', '19', '20', '21');
    var targetUseCaseArray = new Array('C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R');
  
  
  
    //process sum data into targetsheet
  
    var workingSum = 0;
    var solution = 0;
    var usecase = 0;

      while(solution < targetSolutionsArray.length) {

       for(var i=1; i<=dataSheet.getLastRow(); i++){

         if(dataSheet.getRange(solutionsArray[solution] + i).getValue() == 1) {
            
        
            while(usecase < targetUseCaseArray.length) {
              
              if(dataSheet.getRange(useCaseArray[usecase] + i).getValue() == 1) {
                
                targetSheet.getRange(targetUseCaseArray[usecase] + targetSolutionsArray[solution]).setValue(
                  targetSheet.getRange(targetUseCaseArray[usecase] + targetSolutionsArray[solution]).getValue() + 
                  dataSheet.getRange('O' + i).getValue()
                );
              }                  
              usecase++;
            }
         }
         usecase = 0;
       }
       solution++;
       }
  
}
