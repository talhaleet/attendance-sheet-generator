let currentTableData = null;

    document.getElementById('attendance-form').addEventListener('submit', function (e) {
      e.preventDefault();

      const monthInput = document.getElementById('month').value;
      const className = document.getElementById('className').value;
      const namesRaw = document.getElementById('bulkNames').value.trim();
      const fathersRaw = document.getElementById('bulkFathers').value.trim();
      const studentNames = namesRaw.split('\n');
      const fatherNames = fathersRaw.split('\n');
      const [year, month] = monthInput.split('-');
      const daysInMonth = new Date(year, month, 0).getDate();

      const table = document.createElement('table');
      const thead = document.createElement('thead');

      const headerRow1 = document.createElement('tr');
      const th = document.createElement('th');
      th.colSpan = 3 + daysInMonth;
      th.innerText = `${className} - ${new Date(year, month - 1).toLocaleString('default', { month: 'long' })} ${year}`;
      th.classList.add('text-center');
      headerRow1.appendChild(th);
      thead.appendChild(headerRow1);

      const headerRow2 = document.createElement('tr');
      headerRow2.innerHTML = '<th>Sr #</th><th style="min-width: 120px">Student\'s Name</th><th style="min-width: 120px">Father Name</th>';

      let sundayIndexes = [];
      for (let i = 1; i <= daysInMonth; i++) {
        const date = new Date(year, month - 1, i);
        if (date.getDay() === 0) sundayIndexes.push(i);
        const th = document.createElement('th');
        th.innerText = i;
        headerRow2.appendChild(th);
      }

      thead.appendChild(headerRow2);
      table.appendChild(thead);

      const tbody = document.createElement('tbody');
      let sundayRendered = {};
      for (let i = 0; i < studentNames.length; i++) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${i + 1}</td><td>${studentNames[i]}</td><td>${fatherNames[i]}</td>`;
        for (let j = 1; j <= daysInMonth; j++) {
          const td = document.createElement('td');
          if (sundayIndexes.includes(j) && !sundayRendered[j]) {
            td.rowSpan = studentNames.length;
            td.innerHTML = "S<br>U<br>N<br>D<br>A<br>Y";
            td.style.verticalAlign = "middle";
            td.style.textAlign = "center";
            td.style.border = "1px solid #ccc";
            tr.appendChild(td);
            sundayRendered[j] = true;
          } else if (!sundayIndexes.includes(j)) {
            tr.appendChild(td);
          }
        }
        tbody.appendChild(tr);
      }

      table.appendChild(tbody);
      const container = document.getElementById('sheet-container');
      container.innerHTML = '';
      container.appendChild(table);

      currentTableData = {
        className,
        month: new Date(year, month - 1).toLocaleString('default', { month: 'long' }),
        year,
        daysInMonth,
        studentNames,
        fatherNames,
        sundayIndexes
      };

      document.getElementById('excel-btn').disabled = false;
    });

    function downloadExcel() {
      if (!currentTableData) return;
      
      const wb = XLSX.utils.book_new();
      const wsData = [];
      
      // Add title row
      wsData.push([`${currentTableData.className} - ${currentTableData.month} ${currentTableData.year}`]);
      
      // Add header row
      const headers = ['Sr #', 'Student\'s Name', 'Father Name'];
      for (let i = 1; i <= currentTableData.daysInMonth; i++) {
        headers.push(i);
      }
      wsData.push(headers);
      
      // Add student data
      for (let i = 0; i < currentTableData.studentNames.length; i++) {
        const row = [
          i + 1,
          currentTableData.studentNames[i],
          currentTableData.fatherNames[i]
        ];
        
        for (let j = 1; j <= currentTableData.daysInMonth; j++) {
          if (currentTableData.sundayIndexes.includes(j)) {
            // For Sunday columns, add "SUN" only for first row
            row.push(i === 0 ? "SUN" : "");
          } else {
            row.push("");
          }
        }
        wsData.push(row);
      }
      
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      
      // Merge Sunday cells
      currentTableData.sundayIndexes.forEach(day => {
        const colIndex = 3 + day - 1; // 3 initial columns + day - 1 (0-based)
        const range = {
          s: { r: 1, c: colIndex }, // header row
          e: { r: 1 + currentTableData.studentNames.length, c: colIndex }
        };
        if (!ws['!merges']) ws['!merges'] = [];
        ws['!merges'].push(range);
      });
      
      // Merge title row
      const titleRange = {
        s: { r: 0, c: 0 },
        e: { r: 0, c: 3 + currentTableData.daysInMonth - 1 }
      };
      ws['!merges'].push(titleRange);
      
      XLSX.utils.book_append_sheet(wb, ws, "Attendance");
      XLSX.writeFile(wb, `${currentTableData.className}_${currentTableData.month}_Attendance.xlsx`);
    }