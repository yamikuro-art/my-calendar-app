document.getElementById('show-calendar-btn').addEventListener('click', function() {
    const year = document.getElementById('year').value;
    const month = document.getElementById('month').value;
    generateCalendar(year, month);
});

document.getElementById('save-btn').addEventListener('click', function() {
    const year = document.getElementById('year').value;
    const month = document.getElementById('month').value;
    saveData(year, month);
});

function generateCalendar(year, month) {
    const calendarContainer = document.getElementById('calendar-container');
    calendarContainer.innerHTML = ''; // Clear previous calendar

    const firstDay = new Date(year, month - 1, 1);
    const daysInMonth = new Date(year, month, 0).getDate();

    const daysOfWeek = ['日', '月', '火', '水', '木', '金', '土'];
    daysOfWeek.forEach(day => {
        const cell = document.createElement('div');
        cell.className = 'calendar-cell';
        cell.textContent = day;
        calendarContainer.appendChild(cell);
    });

    for (let i = 0; i < firstDay.getDay(); i++) {
        const emptyCell = document.createElement('div');
        emptyCell.className = 'calendar-cell empty';
        calendarContainer.appendChild(emptyCell);
    }

    for (let day = 1; day <= daysInMonth; day++) {
        const cell = document.createElement('div');
        cell.className = 'calendar-cell';
        cell.textContent = day;
        const selectBox = document.createElement('select');
        selectBox.className = 'select-box';
        selectBox.appendChild(new Option("", ""));
        ["C1", "C2", "B", "A"].forEach(optionText => {
            const option = document.createElement('option');
            option.value = optionText;
            option.textContent = optionText;
            selectBox.appendChild(option);
        });
        cell.appendChild(selectBox);
        cell.addEventListener('click', function(event) {
            if (event.target === this) {
                this.classList.toggle('holiday');
            }
        });
        calendarContainer.appendChild(cell);
    }
}

function saveData(year, month) {
    const days = document.querySelectorAll('.calendar-cell:not(.empty)');
    const dataToSave = Array.from(days).map(day => {
        const date = parseInt(day.textContent.trim(), 10);  // 数値として日付を取得
        const isHoliday = day.classList.contains('holiday') ? 'Yes' : 'No';
        const selectBox = day.querySelector('.select-box');
        const selectedOption = selectBox ? selectBox.value : ''; // selectBoxが存在するかをチェック
        return { year, month, date, isHoliday, selectedOption };
    });

    console.log('Data to save:', dataToSave); // デバッグ用にコンソールにデータを出力

    fetch('/save_to_excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(dataToSave)
    })
    .then(response => response.json())
    .then(result => {
        if (result.success) {
            alert('データがExcelに正常に保存されました！');
        } else {
            alert('データのExcelへの保存に失敗しました！');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('保存中にエラーが発生しました。');
    });
}
