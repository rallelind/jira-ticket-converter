Getting started

Please run the following command:
```
npm install
````
When everything has finished installing go to index.js file and edit the FILE variable to the path of your excel file.
When that is done you will at the bottom see the function ```convertToTicketFromFile(FILE)``` function.
This function takes in two paramaters: a file path and a sheet which has a default of 1. In the sheet paramater you can
enter the name of a sheet from the given excel file. Thereafter when ```node index.js```is run it will generate a file
inside a folder named csv-output where you should see a CSV file with the name of the sheet provided.