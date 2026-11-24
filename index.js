document.addEventListener("DOMContentLoaded", () => {
    const rentTableBody = document.querySelector("#rent-table tbody");
    const setInitialMonthButton = document.getElementById("set-initial-month");
    const addMonthButton = document.getElementById("add-month");
    const deleteLastRowButton = document.getElementById("delete-last-row");
    const exportButton = document.getElementById("export-excel");

    let lastMonth = null;
    let lastFinalValue = null;

    function formatMonthYear(date) {
        const options = { year: 'numeric', month: 'long' };
        return new Intl.DateTimeFormat('es-ES', options).format(date);
    }

    setInitialMonthButton.addEventListener("click", () => {
        const initialMonthInput = document.getElementById("initial-month").value;
        const initialValue = parseFloat(document.getElementById("initial-value").value);

        if (!initialMonthInput || isNaN(initialValue)) {
            alert("Por favor, ingrese un mes y un valor inicial.");
            return;
        }

        const [year, month] = initialMonthInput.split("-");
        const baseDate = new Date(year, month - 1); // Mes ingresado
        const previousMonth = new Date(year, month - 2); // Mes anterior

        rentTableBody.innerHTML = "";
        addRow(previousMonth, initialValue, 0);
    });

    addMonthButton.addEventListener("click", () => {
        const inflationRate = parseFloat(document.getElementById("inflation-rate").value);
        if (isNaN(inflationRate)) {
            alert("Ingrese un porcentaje de inflación.");
            return;
        }

        if (!lastMonth || lastFinalValue === null) {
            alert("Primero debe ingresar el mes inicial.");
            return;
        }

        const nextMonth = new Date(lastMonth.getFullYear(), lastMonth.getMonth() + 1);
        addRow(nextMonth, lastFinalValue, inflationRate);
    });

    function addRow(date, initialValue, inflationRate) {
        const increment = (initialValue * inflationRate) / 100;
        const finalValue = initialValue + increment;

        const row = document.createElement("tr");
        row.innerHTML = `
            <td>${formatMonthYear(date)}</td>
            <td>${initialValue.toFixed(2)}</td>
            <td>${inflationRate.toFixed(2)}</td>
            <td>${increment.toFixed(2)}</td>
            <td>${finalValue.toFixed(2)}</td>
        `;
        rentTableBody.appendChild(row);

        lastMonth = date;
        lastFinalValue = finalValue;
    }

    deleteLastRowButton.addEventListener("click", () => {
        if (rentTableBody.rows.length > 0) {
            rentTableBody.removeChild(rentTableBody.lastChild);

            if (rentTableBody.rows.length > 0) {
                const lastRow = rentTableBody.lastElementChild;
                lastMonth = new Date(lastRow.cells[0].textContent + " 01");
                lastFinalValue = parseFloat(lastRow.cells[4].textContent);
            } else {
                lastMonth = null;
                lastFinalValue = null;
            }
        }
    });

    exportButton.addEventListener("click", () => {
        let table = document.getElementById("rent-table");
        let ws = XLSX.utils.table_to_sheet(table, { raw: true });

        // Asegurar que la primera columna (mes/año) se guarde correctamente
        for (let i = 2; i <= rentTableBody.rows.length + 1; i++) {
            let cell = ws[`A${i}`]; // A2, A3, etc.
            if (cell) {
                let dateParts = cell.v.split(" ");
                if (dateParts.length === 2) {
                    let month = dateParts[0];
                    let year = dateParts[1];
                    cell.v = `${month} ${year}`;
                }
            }
        }

        let wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Cálculo Inflación");
        XLSX.writeFile(wb, "calculo_inflacion.xlsx");
    });
});