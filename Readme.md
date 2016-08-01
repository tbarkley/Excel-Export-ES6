# Excel-Export-ES6 #
A simple and fast node.js module for exporting data set to Excel xlsx file. Now completely asynchronous!

- Returns the path of the file (in the temp folder) instead of a buffer.
- Written in ES6.  this helps to free RAM.
- Row can also be a Stream of Array (row).

## Using excel-export ##
Setup configuration object before passing it into the execute method.  **cols** is an array for column definition.  Column definition should have caption and type properties while width property is not required.  The unit for width property is character.   **beforeCellWrite** callback is optional.  beforeCellWrite is invoked with row, cell data and option object (eOpt detail later) parameters.  The return value from beforeCellWrite is what will get written into the cell.  

    let callback = function (error, path) {
        //whatever you want
    };

    let configuration = {
        cols: columns, // Array that defines each columns
        stylesXmlFile: "PathToStyleSheet", //Optional: appends styleSheet xml to Excel
        rows: rows, // Data to be written
        name: "Name of the Tab"
    };

    ExcelExport.execute(configuration, callback);
    
**OR**

    // Array of configurations for multiple pages
    let configuration = [
        configuration,
        configuration
    ];

    ExcelExport.execute(configuration, callback);

## Supported types ##
Supported valid types are;

- string
- date 
- bool
- number

**rows** is the data to be exported. It is an Array of Array (row) or a Stream of Arrays (row). Each row should be the same length as cols.

## Styling ##
Styling is optional.  However, if you want to style your spreadsheet, a valid excel styles xml file is needed.  An easy way to get a styles xml file is to unzip an existing xlsx file which has the desired styles and copy out the styles.xml file. Use **stylesXmlFile** property of configuartion object to specify the relative path and file name of the xml file.  Google for "spreadsheetml style" to learn more detail on styling spreadsheet.  eOpt in beforeCellWrite callback contains rowNum for current row number. eOpt.styleIndex should be a valid zero based index from cellXfs tag of the selected styles xml file.  eOpt.cellType is default to the type value specified in column definition.  However, in some scenario you might want to change it for different format. 

### Basic Example ###
**Using Array of Arrays for Rows and Single Config Object**

    let columns = [{
        caption: 'Title 1',
        type: 'string'
        width: 28.7109375
    }, {
        caption: 'Title 2',
        type: 'date'
    }, {
        caption: 'bool',
        type: 'bool'
    }];
    
    let rows = [
        ['row 1 column 1', 'row 1 column 2'],
        ['row 2 column 1', 'row 2 column 2']
    ];
    
    //since a single config is used, only one sheet will be generated
    let conf = {
        cols: columns,
        stylesXmlFile: "PathToStyleSheet",
        name: "Name of the Tab",
        rows: rows
    };

    return nodeExcel.execute(conf, function(err, path) {
        console.log("Path to excel file", path);
    });
    
### Advanced Example ###
**Using Array of Mongoose Cursors for Rows and Array of Config Objects**

    app.get('/Excel/multisheet', (req, res) => {

        //get cursor for inactive users
        let _getDeactivatedUsersCursor = () => {
        
            return new Promise((resolve, reject) => {
            
                let inactiveUsersQuery = {
                    active: false
                };
                
                //use mongoose to query user collection
                let inactiveUsers = User
                    .find(inactiveUsersQuery)
                    .lean()
                    .limit(100)
                    .cursor(); //get cursor

                inactiveUsers.active = false;
                
                resolve(inactiveUsers);
            });
        };

        //get cursor for active users
        let _getActiveUsersCursor = () => {
        
            return new Promise((resolve, reject) => {
            
                let activeUsersQuery = {
                    active: true
                };
                
                let activeUsers = User
                    .find(activeUsersQuery)
                    .lean()
                    .limit(100)
                    .cursor(); //get cursor
                activeUsers.active = true;
                
                resolve(activeUsers);
            });
        };

        //columns for the report
        let _getReportColumns = () => {
        
            return [{
                beforeCellWrite : (row, cellData, eOpt) => {
                    //Do something to transform row, cellData or cell style
                },
                styleIndex      : 2, //Index from stylesXmlFile
                caption         : 'Deactivation Date',
                type            : 'date',
            }, {
                styleIndex: 2,
                caption   : 'Activation Date',
                type      : 'date'
            }, {
                styleIndex: 2,
                caption   : 'First Name',
                type      : 'string',
            }, {
                styleIndex: 2,
                caption    : 'Last Name',
                type       : 'string',
            }];
        };

        //get report stream
        let _getReportStream = (cursor) => {
        
            let userStream = new stream.Transform({objectMode: true});
            
            //Implement _transform method
            userStream._transform = (user, encoding, callback) => {
            
                callback(null, [{
                    deactivationDate: user.deactivationDate,
                    activationDate  : user.activationDate,
                    firstName       : user.firstName,
                    lastName        : user.lastName
                }]);
            };
            
            //Pipe cursor to stream
            return cursor.pipe(userStream);
        };

        Promise.all([
            _getActiveUsersCursor, 
            _getDeactivatedUsersCursor
        ]).then((usersCursor) => {
        
            //array to hold config objects
            let configs = [];
            let config = {
                stylesXmlFile: REPORT_STYLES_PATH,
                name         : 'No Users',
                cols         : _getReportColumns(),
                rows         : []
            };

            //create array of config objects
            configs = usersCursor.map((cursor) => {
                config.rows = _getReportStream(cursor);
                if (cursor.active == true) {
                    config.name = "Active Users";
                    config.cols = _getReportColumns();
                } else if (cursor.active == false) {
                   config.name = "Deactivated Users";
                   config.cols = _getReportColumns();
                }
                return config;
            });

            //since array of configs is passed in, resulting xlsx file will have multiple sheets (tabs)
            return nodeExcel.execute(configs, (err, path) => {
                res.sendFile(path);
            });
        }).catch((err) => {
            //log error
            console.error(err);
            res.sendStatus(500);
        });
    });

    app.listen(3000);
    console.log('Listening on port 3000');
