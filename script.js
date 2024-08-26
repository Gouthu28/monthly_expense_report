let rowCount = 0;

function addRow() {
    rowCount++;
    const tbody = document.getElementById('expenseBody');
    const row = document.createElement('tr');

    row.innerHTML = `
        <td>${rowCount}</td>
        <td><input type="date" onchange="handleDateChange(this)"></td>
        <td><input type="text" placeholder="Area of Work"></td>
        <td>
            <select>
                <option value="Bike">Bike</option>
                <option value="Bus">Bus</option>
                <option value="Auto">Auto</option>
                <option value="Taxi">Taxi</option>
                <option value="Train">Train</option>
            </select>
        </td>
        <td><input type="number" placeholder="Km" oninput="calculateRow(${rowCount})"></td>
        <td><input type="number" placeholder="Travel Bill" oninput="calculateRow(${rowCount})"></td>
        <td><input type="number" value="100" oninput="calculateRow(${rowCount})"></td>
        <td><input type="number" placeholder="Petrol Bill" readonly></td>
        <td><input type="number" placeholder="Hotel Bill" oninput="calculateRow(${rowCount})"></td>
        <td><input type="number" placeholder="Other Expenses" oninput="calculateRow(${rowCount})"></td>
        <td><input type="number" placeholder="Total Expense" readonly></td>
        <td><input type="text" placeholder="Remark"></td>
    `;

    tbody.appendChild(row);
}

function handleDateChange(input) {
    const dateValue = input.value;
    const rows = document.querySelectorAll('#expenseBody tr');
    let allowanceAdded = false;

    rows.forEach(row => {
        const dateInput = row.children[1].children[0].value;
        if (dateInput === dateValue) {
            if (!allowanceAdded) {
                row.children[6].children[0].value = 100; // Add allowance once
                allowanceAdded = true;
            } else {
                row.children[6].children[0].value = 0; // No allowance for subsequent entries of the same date
            }
            calculateRow(parseInt(row.children[0].innerText));
        }
    });
}

function calculateRow(rowId) {
    const row = document.querySelectorAll(`#expenseBody tr`)[rowId - 1];
    const kmTravelled = parseFloat(row.children[4].children[0].value) || 0;
    const travelBill = parseFloat(row.children[5].children[0].value) || 0;
    const allowance = parseFloat(row.children[6].children[0].value) || 0;
    const hotelBill = parseFloat(row.children[8].children[0].value) || 0;
    const otherExpenses = parseFloat(row.children[9].children[0].value) || 0;

    // Constants
    const fuelPrice = 105.88;
    const vehicleMileage = 45;

    // Calculate petrol bill
    const petrolBill = (kmTravelled / vehicleMileage) * fuelPrice;
    row.children[7].children[0].value = petrolBill.toFixed(2);

    // Calculate total expense for the row
    const totalExpense = travelBill + allowance + petrolBill + hotelBill + otherExpenses;
    row.children[10].children[0].value = totalExpense.toFixed(2);

    calculateGrandTotal();
}

function calculateGrandTotal() {
    let grandTotal = 0;
    document.querySelectorAll('#expenseBody tr').forEach(row => {
        grandTotal += parseFloat(row.children[10].children[0].value) || 0;
    });
    document.getElementById('grandTotal').innerText = `$${grandTotal.toFixed(2)}`;
}

function exportToExcel() {
    const table = document.getElementById("expenseTable");
    const rows = [];
    const mergedRows = {};

    // Collect data and merge rows with the same date
    for (let i = 1, row; row = table.rows[i]; i++) {
        const date = row.cells[1].children[0].value;
        const data = {
            areaOfWork: row.cells[2].children[0].value,
            modeOfTravel: row.cells[3].children[0].value,
            kmTravelled: parseFloat(row.cells[4].children[0].value) || 0,
            travelBill: parseFloat(row.cells[5].children[0].value) || 0,
            allowance: parseFloat(row.cells[6].children[0].value) || 0,
            petrolBill: parseFloat(row.cells[7].children[0].value) || 0,
            hotelBill: parseFloat(row.cells[8].children[0].value) || 0,
            otherExpenses: parseFloat(row.cells[9].children[0].value) || 0,
            totalExpense: parseFloat(row.cells[10].children[0].value) || 0,
            remark: row.cells[11].children[0].value || ''
        };

        if (mergedRows[date]) {
            mergedRows[date].kmTravelled += data.kmTravelled;
            mergedRows[date].travelBill += data.travelBill;
            mergedRows[date].allowance = Math.max(mergedRows[date].allowance, data.allowance);
            mergedRows[date].petrolBill += data.petrolBill;
            mergedRows[date].hotelBill += data.hotelBill;
            mergedRows[date].otherExpenses += data.otherExpenses;
            mergedRows[date].totalExpense += data.totalExpense;
            mergedRows[date].remark += ' | ' + data.remark;
        } else {
            mergedRows[date] = data;
        }
    }

    // Create rows for CSV
    for (const date in mergedRows) {
        const row = [
            date,
            mergedRows[date].areaOfWork,
            mergedRows[date].modeOfTravel,
            mergedRows[date].kmTravelled.toFixed(2),
            mergedRows[date].travelBill.toFixed(2),
            mergedRows[date].allowance.toFixed(2),
            mergedRows[date].petrolBill.toFixed(2),
            mergedRows[date].hotelBill.toFixed(2),
            mergedRows[date].otherExpenses.toFixed(2),
            mergedRows[date].totalExpense.toFixed(2),
            mergedRows[date].remark
        ];
        rows.push(row.join(","));
    }

    let csvContent = "Date,Area of Work,Mode of Travel,Kilometers Traveled,Travel Bill,Allowance,Petrol Bill,Hotel Bill,Other Expenses,Total Expense,Remark\n";
    csvContent += rows.join("\n");

    let blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    let url = URL.createObjectURL(blob);
    let a = document.createElement("a");
    a.href = url;
    a.download = "monthly_expense_report.csv";
    a.click();
}
