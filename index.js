let currentTableData = null;

    document.getElementById('month').addEventListener('change', function() {
      const monthInput = this.value;
      if (!monthInput) return;
      
      const [year, month] = monthInput.split('-');
      const daysInMonth = new Date(year, month, 0).getDate();
      
      const startDaySelect = document.getElementById('startDay');
      const endDaySelect = document.getElementById('endDay');
      
      // Clear existing options
      startDaySelect.innerHTML = '';
      endDaySelect.innerHTML = '';
      
      // Add options for each day
      for (let i = 1; i <= daysInMonth; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        startDaySelect.appendChild(option.cloneNode(true));
        endDaySelect.appendChild(option);
      }
      
      // Set default values
      startDaySelect.value = 1;
      endDaySelect.value = daysInMonth;
      
      // Show the day range container
      document.getElementById('day-range-container').style.display = 'block';
    });

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
      const studentCount = parseInt(document.getElementById('studentCount').value);
      const startDay = parseInt(document.getElementById('startDay').value);
      const endDay = parseInt(document.getElementById('endDay').value);
      
      // Validate day range
      if (startDay > endDay) {
        alert('Start day cannot be greater than end day');
        return;
      }
      
      const table = document.createElement('table');
      const thead = document.createElement('thead');

      // Pad missing entries if fewer than studentCount
       while (studentNames.length < studentCount) studentNames.push("");
       while (fatherNames.length < studentCount) fatherNames.push("");

      const headerRow1 = document.createElement('tr');
      const th = document.createElement('th');
      th.colSpan = 3 + (endDay - startDay + 1);
      th.innerText = `${className} - ${new Date(year, month - 1).toLocaleString('default', { month: 'long' })} ${year} (Days ${startDay}-${endDay})`;
      th.classList.add('text-center');
      headerRow1.appendChild(th);
      thead.appendChild(headerRow1);

      const headerRow2 = document.createElement('tr');
      headerRow2.innerHTML = '<th>Sr #</th><th style="min-width: 120px">Student\'s Name</th><th style="min-width: 120px">Father Name</th>';

      let sundayIndexes = [];
      for (let i = startDay; i <= endDay; i++) {
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
      for (let i = 0; i < studentCount; i++) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${i + 1}</td><td>${studentNames[i]}</td><td>${fatherNames[i]}</td>`;
        for (let j = startDay; j <= endDay; j++) {
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
        sundayIndexes,
        startDay,
        endDay
      };

      document.getElementById('excel-btn').disabled = false;
    });

    function downloadExcel() {
      if (!currentTableData) return;
      
      const wb = XLSX.utils.book_new();
      const wsData = [];
      
      // Add title row
      wsData.push([`${currentTableData.className} - ${currentTableData.month} ${currentTableData.year} (Days ${currentTableData.startDay}-${currentTableData.endDay})`]);
      
      // Add header row
      const headers = ['Sr #', 'Student\'s Name', 'Father Name'];
      for (let i = currentTableData.startDay; i <= currentTableData.endDay; i++) {
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
        
        for (let j = currentTableData.startDay; j <= currentTableData.endDay; j++) {
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
        const colIndex = 3 + day - currentTableData.startDay; // 3 initial columns + day - startDay (0-based)
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
        e: { r: 0, c: 3 + (currentTableData.endDay - currentTableData.startDay) }
      };
      ws['!merges'].push(titleRange);
      
      XLSX.utils.book_append_sheet(wb, ws, "Attendance");
      XLSX.writeFile(wb, `${currentTableData.className}_${currentTableData.month}_Attendance_${currentTableData.startDay}-${currentTableData.endDay}.xlsx`);
    }