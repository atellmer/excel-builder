require.config({
                text: 'text.js',
                paths: {
                    JSZip: '../../jszip.min'
                },
                shim: {
                    'JSZip': {
                        exports: 'JSZip'
                    }
                }
            });
            require(['../excel-builder', 'text!testdata.json', '../Template/BasicReport'], function (builder, testdata, BasicReport) {
                var data = JSON.parse(testdata);
                /* for converting data sets from key-> value to array of values
                var stuff = [];
                for(var i = 0, l = data.length; i < l; i++) {
                    var d = data[i];
                    stuff[i] = [
                        d.id,
                        d.name,
                        d.price,
                        d.location,
                        d.startDate,
                        d.endDate
                    ];
                }
                */
                var basicReport = new BasicReport();
                var columns = [
                    {id: 'id', name: "ID", type: 'number', width: 20},
                    {id: 'name', name:"Name", type: 'string', width: 50},
                    {id: 'price', name: "Price", type: 'number', style: basicReport.predefinedFormatters.currency.id},
                    {id: 'location', name: "Location", type: 'string'},
                    {id: 'startDate', name: "Start Date", type: 'date', style: basicReport.predefinedFormatters.date.id, width: 15},
                    {id: 'endDate', name: "End Date", type: 'date', style: basicReport.predefinedFormatters.date.id, width: 15}
                ];
                
                var worksheetData = [
                    [
                        {value: "ID", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}, 
                        {value: "Name", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}, 
                        {value: "Price", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}, 
                        {value: "Location", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}, 
                        {value: "Start Date", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
                        {value: "End Date", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}
                    ]
                ].concat(data);
                
                basicReport.setHeader([
                    {bold: true, text: "Generic Report"}, "", ""
                ]);
                basicReport.setData(worksheetData);
                basicReport.setColumns(columns);
                basicReport.setFooter([
                    '', '', 'Page &P of &N'
                ]);
                if('download' in document.createElement('a')){
                    $("#downloader").attr({
                        href: "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"+builder.createFile(basicReport.prepare())
                    });
                } else {
                    Downloadify.create('downloader',{
                            filename: function(){
                                    return "sample.xlsx";
                            },
                            data: function(){ 
                                    return builder.createFile(basicReport.prepare());
                            },
//                            onComplete: function(){ alert('Your File Has Been Saved!'); },
//                            onCancel: function(){ alert('You have cancelled the saving of this file.'); },
//                            onError: function(){ alert('You must put something in the File Contents or there will be nothing to save!'); },
                            swf: 'downloadify/media/downloadify.swf',
                            downloadImage: 'downloadify/images/download.png',
                            width: 100,
                            dataType: 'base64',
                            height: 30,
                            transparent: true,
                            append: false
                    });
                }
            });