google.charts.load('current', {'packages':['bar']});
      google.charts.setOnLoadCallback(drawStuff);

      function drawStuff() {
        // Log variables to check if they are being fetched correctly
       

        var data = new google.visualization.arrayToDataTable([
          ['Course Outcome', 'Level'],
          ["ie1-1", c_ie1_1],
          ["ie1-2", c_ie1_2],
          ["mte1", c_mte_1],
          ["mte2", c_mte_2],
          ['mte3', c_mte_3],
          ["ie2-1", c_ie2_1],
          ["ie2-2", c_ie2_2],
          ["ete1", c_ete_1],
          ["ete2", c_ete_2],
          ["ete3", c_ete_3],
          ["ete4", c_ete_4],
          ["ete5", c_ete_5],
          ["ete6", c_ete_6],
        ]);

        var options = {
          width: 800,
          legend: { position: 'none' },
          axes: {
            x: {
              0: { side: 'top', label: 'Course Outcomes'} // Top x-axis.
            }
          },
          bar: { groupWidth: "90%" }
        };

        var chart = new google.charts.Bar(document.getElementById('top_x_div'));
        // Convert the Classic options to Material options.
        chart.draw(data, google.charts.Bar.convertOptions(options));
      }