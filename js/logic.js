
  var c_ie1_1 = 0 ;
  var c_ie1_2 = 0 ;
  var c_mte_1 = 0 ;
  var c_mte_2 = 0 ;
  var c_mte_3 = 0 ;
  var c_ie2_1 = 0 ;
  var c_ie2_2 = 0 ;
  var c_ete_1 = 0 ;
  var c_ete_2 = 0 ;
  var c_ete_6 = 0 ;
  var c_ete_3 = 0 ;
  var c_ete_4 = 0 ;
  var c_ete_5 = 0 ;

const excel_file = document.getElementById('excel_file');

excel_file.addEventListener('change', (event) => 
    {

     if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(event.target.files[0].type)) {
          document.getElementById('excel_data').innerHTML = '<div class="alert alert-danger" >Only .xlsx or .xls file format are allowed</div>';

          excel_file.value = '';

          return false;
     }

     var reader = new FileReader();

     reader.readAsArrayBuffer(event.target.files[0]);

     reader.onload = function (event) {

          var data = new Uint8Array(reader.result);
                    
          var work_book = XLSX.read(data, { type: 'array' });

          var sheet_name = work_book.SheetNames;

          var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], { header: 1 });
          total_number = sheet_data.length - 8;
          



          if (sheet_data.length > 0) {
               var table_output = '<table class="table table-striped table-bordered">';
                    var qualify_ie1_1 = 0;
                    var qualify_ie1_2 = 0;

                    var qualify_mte_1 = 0;
                    var qualify_mte_2 = 0;
                    var qualify_mte_3 = 0;

                    var qualify_ie2_1 = 0;
                    var qualify_ie2_2 = 0;

                    var qualify_ete_1 = 0;
                    var qualify_ete_2 = 0;
                    var qualify_ete_3 = 0;
                    var qualify_ete_4 = 0;
                    var qualify_ete_5 = 0;
                    var qualify_ete_6 = 0;

               var max_ie1_1 = sheet_data[1][4];
               var max_ie1_2 = sheet_data[1][5];
               var max_mte_1 = sheet_data[1][6];
               var max_mte_2 = sheet_data[1][7];
               var max_mte_3 = sheet_data[1][8];
               var max_ie2_1 = sheet_data[1][9];
               var max_ie2_2 = sheet_data[1][10];
               var max_ete_1 = sheet_data[1][11];
               var max_ete_2 = sheet_data[1][12];
               var max_ete_3 = sheet_data[1][13];
               var max_ete_4 = sheet_data[1][14];
               var max_ete_5 = sheet_data[1][15];
               var max_ete_6 = sheet_data[1][16];

               for (var row = 4; row < sheet_data.length - 4; row++) {
                    
                         for (var cell = 4; cell <= 4; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ie1_1 = sheet_data[row][cell];
                              }
                              else {
                                   var ie1_1 = 0;
                              }
                         }
                         for (var cell = 5; cell <= 5; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ie1_2 = sheet_data[row][cell];
                              }
                              else {
                                   var ie1_2 = 0;
                              }
                         }

                         for (var cell = 6; cell <= 6; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var mte_1 = sheet_data[row][cell];
                              }
                              else {
                                   var mte_1 = 0;
                              }
                         }
                         for (var cell = 7; cell <= 7; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var mte_2 = sheet_data[row][cell];
                              }
                              else {
                                   var mte_2 = 0;
                              }
                         }
                         for (var cell = 8; cell <= 8; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var mte_3 = sheet_data[row][cell];
                              }
                              else {
                                   var mte_3 = 0;
                              }
                         }

                         for (var cell = 9; cell <= 9; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ie2_1 = sheet_data[row][cell];
                              }
                              else {
                                   var ie2_1 = 0;
                              }
                         }
                         for (var cell = 10; cell <= 10; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ie2_2 = sheet_data[row][cell];
                              }
                              else {
                                   var ie2_2 = 0;
                              }
                         }

                         for (var cell = 11; cell <= 11; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ete_1 = sheet_data[row][cell];
                              }
                              else {
                                   var ete_1 = 0;
                              }
                         }
                         for (var cell = 12; cell <= 12; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ete_2 = sheet_data[row][cell];
                              }
                              else {
                                   var ete_2 = 0;
                              }
                         }
                         for (var cell = 13; cell <= 13; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ete_3 = sheet_data[row][cell];
                              }
                              else {
                                   var ete_3 = 0;
                              }
                         }
                         for (var cell = 14; cell <= 14; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ete_4 = sheet_data[row][cell];
                              }
                              else {
                                   var ete_4 = 0;
                              }
                         }
                         for (var cell = 15; cell <= 15; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ete_5 = sheet_data[row][cell];
                              }
                              else {
                                   var ete_5 = 0;
                              }
                         }
                         for (var cell = 16; cell <= 16; cell++) {
                              if (sheet_data[row][cell] >= 0 || sheet_data[row][cell] < 0) {
                                   var ete_6 = sheet_data[row][cell];
                              }
                              else {
                                   var ete_6 = 0;
                              }
                         }

                         var ie1_p_1 = (ie1_1) /(max_ie1_1) * 100;
                         var ie1_p_2 = (ie1_2) /(max_ie1_2) * 100;

                         var mte_p_1 = (mte_1) / (max_mte_1) * 100;
                         var mte_p_2 = (mte_2) / (max_mte_2) * 100;
                         var mte_p_3 = (mte_3) / (max_mte_3) * 100;

                         var ie2_p_1 = (ie2_1) / (max_ie2_1) * 100;
                         var ie2_p_2 = (ie2_2) / (max_ie2_2) * 100;

                         var ete_p_1 = (ete_1) / (max_ete_1) * 100;
                         var ete_p_2 = (ete_2) / (max_ete_2) * 100;
                         var ete_p_3 = (ete_3) / (max_ete_3) * 100;
                         var ete_p_4 = (ete_4) / (max_ete_4) * 100;
                         var ete_p_5 = (ete_5) / (max_ete_5) * 100;
                         var ete_p_6 = (ete_6) / (max_ete_6) * 100;

                         if (ie1_p_1 >= 70) {
                              qualify_ie1_1++;
                         }
                         if (ie1_p_2 >= 70) {
                              qualify_ie1_2++;
                         }

                         if (mte_p_1 >= 70) {
                              qualify_mte_1++;
                         }
                         if (mte_p_2 >= 70) {
                              qualify_mte_2++;
                         }
                         if (mte_p_3 >= 70) {
                              qualify_mte_3++;
                         }

                         if (ie2_p_1 >= 70) {
                              qualify_ie2_1++;
                         }
                         if (ie2_p_2 >= 70) {
                              qualify_ie2_2++;
                         }

                         if (ete_p_1 >= 70) {
                              qualify_ete_1++;
                         }
                         if (ete_p_2 >= 70) {
                              qualify_ete_2++;
                         }
                         if (ete_p_3 >= 70) {
                              qualify_ete_3++;
                         }
                         if (ete_p_4 >= 70) {
                              qualify_ete_4++;
                         }
                         if (ete_p_5 >= 70) {
                              qualify_ete_5++;
                         }
                         if (ete_p_6 >= 70) {
                              qualify_ete_6++;
                         }


                    

                    
               }
               var qualify_ie1_p_1 = (qualify_ie1_1 / total_number) * 100;
               var qualify_ie1_p_2 = (qualify_ie1_2 / total_number) * 100;

               var qualify_mte_p_1 = (qualify_mte_1 / total_number) * 100;
               var qualify_mte_p_2 = (qualify_mte_2 / total_number) * 100;
               var qualify_mte_p_3 = (qualify_mte_3 / total_number) * 100;

               var qualify_ie2_p_1 = (qualify_ie2_1 / total_number) * 100;
               var qualify_ie2_p_2 = (qualify_ie2_2 / total_number) * 100;

               var qualify_ete_p_1 = (qualify_ete_1 / total_number) * 100;
               var qualify_ete_p_2 = (qualify_ete_2 / total_number) * 100;
               var qualify_ete_p_3 = (qualify_ete_3 / total_number) * 100;
               var qualify_ete_p_4 = (qualify_ete_4 / total_number) * 100;
               var qualify_ete_p_5 = (qualify_ete_5 / total_number) * 100;
               var qualify_ete_p_6 = (qualify_ete_6 / total_number) * 100;


               var c_ie1_1 = 0;
               var c_ie1_2 = 0;

               var c_mte_1 = 0;
               var c_mte_2 = 0;
               var c_mte_3 = 0;

               var c_ie2_1 = 0;
               var c_ie2_2 = 0;

               var c_ete_1 = 0;
               var c_ete_2 = 0;
               var c_ete_3 = 0;
               var c_ete_4 = 0;
               var c_ete_5 = 0;
               var c_ete_6 = 0;


               if (qualify_ie1_p_1 <= 100 && qualify_ie1_p_1 > 90) {
                    var c_ie1_1 = 3;
               }
               else if (qualify_ie1_p_1 <= 90 && qualify_ie1_p_1 > 80) {
                    var c_ie1_1 = 2;
               }
               else if (qualify_ie1_p_1 <= 80 && qualify_ie1_p_1 > 70) {
                    var c_ie1_1 = 1;
               }

               if (qualify_ie1_p_2 <= 100 && qualify_ie1_p_2 > 90) {
                    var c_ie1_2 = 3;
               }
               else if (qualify_ie1_p_2 <= 90 && qualify_ie1_p_2 > 80) {
                    var c_ie1_2 = 2;
               }
               else if (qualify_ie1_p_2 <= 80 && qualify_ie1_p_2 > 70) {
                    var c_ie1_2 = 1;
               }


               if (qualify_mte_p_1 <= 100 && qualify_mte_p_1 > 90) {
                    var c_mte_1 = 3;
               }
               else if (qualify_mte_p_1 <= 90 && qualify_mte_p_1 > 80) {
                    var c_mte_1 = 2;
               }
               else if (qualify_mte_p_1 <= 80 && qualify_mte_p_1 > 70) {
                    var c_mte_1 = 1;
               }

               if (qualify_mte_p_2 <= 100 && qualify_mte_p_2 > 90) {
                    var c_mte_2 = 3;
               }
               else if (qualify_mte_p_2 <= 90 && qualify_mte_p_2 > 80) {
                    var c_mte_2 = 2;
               }
               else if (qualify_mte_p_2 <= 80 && qualify_mte_p_2 > 70) {
                    var c_mte_2 = 1;
               }

               if (qualify_mte_p_3 <= 100 && qualify_mte_p_3 > 90) {
                    var c_mte_3 = 3;
               }
               else if (qualify_mte_p_3 <= 90 && qualify_mte_p_3 > 80) {
                    var c_mte_3 = 2;
               }
               else if (qualify_mte_p_3 <= 80 && qualify_mte_p_3 > 70) {
                    var c_mte_3 = 1;
               }


               if (qualify_ie2_p_1 <= 100 && qualify_ie2_p_1 > 90) {
                    var c_ie2_1 = 3;
               }
               else if (qualify_ie2_p_1 <= 90 && qualify_ie2_p_1 > 80) {
                    var c_ie2_1 = 2;
               }
               else if (qualify_ie2_p_1 <= 80 && qualify_ie2_p_1 > 70) {
                    var c_ie2_1 = 1;
               }

               if (qualify_ie2_p_2 <= 100 && qualify_ie2_p_2 > 90) {
                    var c_ie2_2 = 3;
               }
               else if (qualify_ie2_p_2 <= 90 && qualify_ie2_p_2 > 80) {
                    var c_ie2_2 = 2;
               }
               else if (qualify_ie2_p_2 <= 80 && qualify_ie2_p_2 > 70) {
                    var c_ie2_2 = 1;
               }


               if (qualify_ete_p_1 <= 100 && qualify_ete_p_1 > 90) {
                    var c_ete_1 = 3;
               }
               else if (qualify_ete_p_1 <= 90 && qualify_ete_p_1 > 80) {
                    var c_ete_1 = 2;
               }
               else if (qualify_ete_p_1 <= 80 && qualify_ete_p_1 > 70) {
                    var c_ete_1 = 1;
               }

               if (qualify_ete_p_2 <= 100 && qualify_ete_p_2 > 90) {
                    var c_ete_2 = 3;
               }
               else if (qualify_ete_p_2 <= 90 && qualify_ete_p_2 > 80) {
                    var c_ete_2 = 2;
               }
               else if (qualify_ete_p_2 <= 80 && qualify_ete_p_2 > 70) {
                    var c_ete_2 = 1;
               }

               if (qualify_ete_p_3 <= 100 && qualify_ete_p_3 > 90) {
                    var c_ete_3 = 3;
               }
               else if (qualify_ete_p_3 <= 90 && qualify_ete_p_3 > 80) {
                    var c_ete_3 = 2;
               }
               else if (qualify_ete_p_3 <= 80 && qualify_ete_p_3 > 70) {
                    var c_ete_3 = 1;
               }

               if (qualify_ete_p_4 <= 100 && qualify_ete_p_4 > 90) {
                    var c_ete_4 = 3;
               }
               else if (qualify_ete_p_4 <= 90 && qualify_ete_p_4 > 80) {
                    var c_ete_4 = 2;
               }
               else if (qualify_ete_p_4 <= 80 && qualify_ete_p_4 > 70) {
                    var c_ete_4 = 1;
               }

               if (qualify_ete_p_5 <= 100 && qualify_ete_p_5 > 90) {
                    var c_ete_5 = 3;
               }
               else if (qualify_ete_p_5 <= 90 && qualify_ete_p_5 > 80) {
                    var c_ete_5 = 2;
               }
               else if (qualify_ete_p_5 <= 80 && qualify_ete_p_5 > 70) {
                    var c_ete_5 = 1;
               }

               if (qualify_ete_p_6 <= 100 && qualify_ete_p_6 > 90) {
                    var c_ete_6 = 3;
               }
               else if (qualify_ete_p_6 <= 90 && qualify_ete_p_6 > 80) {
                    var c_ete_6 = 2;
               }
               else if (qualify_ete_p_6 <= 80 && qualify_ete_p_6 > 70) {
                    var c_ete_6 = 1;
               }


               if (c_ie1_1 == 0) {
                    var max_ie1_1 = 0;
               }

               if (c_ie1_2 == 0) {
                    var max_ie1_2 = 0;
               }


               if (c_mte_1 == 0) {
                    var max_mte_1 = 0;
               }

               if (c_mte_2 == 0) {
                    var max_mte_2 = 0;
               }

               if (c_mte_3 == 0) {
                    var max_mte_3 = 0;
               }


               if (c_ie2_1 == 0) {
                    var max_ie2_1 = 0;
               }

               if (c_ie2_2 == 0) {
                    var max_ie2_2 = 0;
               }


               if (c_ete_1 == 0) {
                    var max_ete_1 = 0;
               }

               if (c_ete_2 == 0) {
                    var max_ete_2 = 0;
               }

               if (c_ete_3 == 0) {
                    var max_ete_3 = 0;
               }

               if (c_ete_4 == 0) {
                    var max_ete_4 = 0;
               }

               if (c_ete_5 == 0) {
                    var max_ete_5 = 0;
               }

               if (c_ete_6 == 0) {
                    var max_ete_6 = 0;
               }


               var wtg_co_1 = ((c_ie1_1 * max_ie1_1) + (c_mte_1 * max_mte_1) + (c_ete_1 * max_ete_1)) / (max_ie1_1 + max_mte_1 + max_ete_1);

               var wtg_co_2 = ((c_ie1_2 * max_ie1_2) + (c_mte_2 * max_mte_2) + (c_ete_2 * max_ete_2)) / (max_ie1_2 + max_mte_2 + max_ete_2);

               var wtg_co_3 = ((c_mte_3 * max_mte_3) + (c_ete_3 * max_ete_3)) / (max_mte_3 + max_ete_3);

               var wtg_co_4 = ((c_ie2_1 * max_ie2_1) + (c_ete_4 * max_ete_4)) / (max_ie2_1 + max_ete_4);

               var wtg_co_5 = ((c_ie2_2 * max_ie2_2) + (c_ete_5 * max_ete_5)) / (max_ie2_2 + max_ete_5);

               var wtg_co_6 = (c_ete_6 * max_ete_6) / (max_ete_6);


               var wtg_qualify_1 = ((qualify_ie1_p_1 * max_ie1_1) + (qualify_mte_p_1 * max_mte_1) + (qualify_ete_p_1 * max_ete_1)) / (max_ie1_1 + max_mte_1 + max_ete_1);

               var wtg_qualify_2 = ((qualify_ie1_p_2 * max_ie1_2) + (qualify_mte_p_2 * max_mte_2) + (qualify_ete_p_2 * max_ete_2)) / (max_ie1_2 + max_mte_2 + max_ete_2);

               var wtg_qualify_3 = ((qualify_mte_p_3 * max_mte_3) + (qualify_ete_p_3 * max_ete_3)) / (max_mte_3 + max_ete_3);

               var wtg_qualify_4 = ((qualify_ie2_p_1 * max_ie2_1) + (qualify_ete_p_4 * max_ete_4)) / (max_ie2_1 + max_ete_4);

               var wtg_qualify_5 = ((qualify_ie2_p_2 * max_ie2_2) + (qualify_ete_p_5 * max_ete_5)) / (max_ie2_2 + max_ete_5);

               var wtg_qualify_6 = (qualify_ete_p_6 * max_ete_6) / (max_ete_6);

            
     
              var headers = [
            '', 'ie1-1', 'ie1-2', 'mte1', 'mte2', 'mte3', 'ie2-1', 'ie2-2', 'ete1', 'ete2', 'ete3', 'ete4', 'ete5', 'ete6'
        ];

        // Define your data for the first table
        var percentages = [
            qualify_ie1_p_1, qualify_ie1_p_2, qualify_mte_p_1, qualify_mte_p_2, qualify_mte_p_3, qualify_ie2_p_1, qualify_ie2_p_2,
            qualify_ete_p_1, qualify_ete_p_2, qualify_ete_p_3, qualify_ete_p_4, qualify_ete_p_5, qualify_ete_p_6
        ];
        var levels = [
            c_ie1_1, c_ie1_2, c_mte_1, c_mte_2, c_mte_3, c_ie2_1, c_ie2_2,
            c_ete_1, c_ete_2, c_ete_3, c_ete_4, c_ete_5, c_ete_6
        ];

        // Define your data for the second table
        var co_percentages = [
            wtg_qualify_1, wtg_qualify_2, wtg_qualify_3, wtg_qualify_4, wtg_qualify_5, wtg_qualify_6
        ];
        var co_levels = [
            wtg_co_1, wtg_co_2, wtg_co_3, wtg_co_4, wtg_co_5, wtg_co_6
        ];

        // Generate the first table HTML
        var table_output_1 = '<h2>Course Outcome Data</h2><table>';

        // Add header row for the first table
        table_output_1 += '<tr><th></th>';
        for (var i = 1; i < headers.length; i++) {
            table_output_1 += '<th>' + headers[i] + '</th>';
        }
        table_output_1 += '</tr>';

        // Add "Course Outcome Percentage" row
        table_output_1 += '<tr><th>Course Outcome Percentage</th>';
        for (var i = 0; i < percentages.length; i++) {
            table_output_1 += '<td>' + percentages[i].toFixed(3) + '%' + '</td>';
        }
        table_output_1 += '</tr>';

        // Add "Course Outcome Level" row
        table_output_1 += '<tr><th>Course Outcome Level</th>';
        for (var i = 0; i < levels.length; i++) {
            table_output_1 += '<td>' + levels[i] + '</td>';
        }
        table_output_1 += '</tr>';

        table_output_1 += '</table>';

        // Generate the second table HTML
        var table_output_2 = '<h2>CO Data</h2><table>';

        // Add header row for the second table
        table_output_2 += '<tr><th>CO</th><th>CO Percentage</th><th>CO Level</th></tr>';

        // Add rows for CO percentages and levels
        for (var i = 0; i < co_percentages.length; i++) {
            table_output_2 += '<tr>';
            table_output_2 += '<th>CO-' + (i + 1) + '</th>';
            table_output_2 += '<td>' + co_percentages[i].toFixed(3) + '%' + '</td>';
            table_output_2 += '<td>' + co_levels[i].toFixed(3) + '</td>';
            table_output_2 += '</tr>';
        }

        table_output_2 += '</table>';

        // Render the tables
        document.getElementById('excel_data').innerHTML = table_output_1 + table_output_2;


                        }
                        excel_file.value = '';
                   }
              
     });

     function updateDataAndRedrawChart() {
        updateData(); // Update the data
        drawStuff(); // Redraw the chart
      }
      
      // Periodically update data and redraw chart
      setInterval(updateDataAndRedrawChart, 5000); // Update and redraw every 5 seconds (adjust as needed)