
document.addEventListener("DOMContentLoaded", () => {
    const rentTableBody = document.querySelector("#rent-table tbody");
    const setInitialMonthButton = document.getElementById("set-initial-month");
    const addMonthButton = document.getElementById("add-month");
    const deleteLastRowButton = document.getElementById("delete-last-row");
    const exportButton = document.getElementById("export-excel");
    const inflationInput = document.getElementById("inflation-rate");

    const API_URL = "https://api.argentinadatos.com/v1/finanzas/indices/inflacion";

    let lastMonth = null;
    let lastFinalValue = null;

    function formatMonthYear(date) {
        const options = { year: "numeric", month: "long" };
        return new Intl.DateTimeFormat("es-AR", options).format(date);
    }

    function formatDateKey(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, "0");
        return `${year}-${month}`;
    }

    setInitialMonthButton.addEventListener("click", () => {
        const initialMonthInput = document.getElementById("initial-month").value;
        const initialValue = Number(document.getElementById("initial-value").value);

        if (!initialMonthInput || !Number.isFinite(initialValue) || initialValue < 0) {
            alert("Por favor, ingrese un mes y un valor inicial válido.");
            return;
        }

        const [year, month] = initialMonthInput.split("-").map(Number);

        const previousMonth = new Date(year, month - 2, 1);

        rentTableBody.innerHTML = "";
        lastMonth = null;
        lastFinalValue = null;

        addRow(previousMonth, initialValue, 0);
    });

    addMonthButton.addEventListener("click", async () => {
        if (!lastMonth || lastFinalValue === null) {
            alert("Primero debe ingresar el mes inicial.");
            return;
        }

        const nextMonth = new Date(
            lastMonth.getFullYear(),
            lastMonth.getMonth() + 1,
            1
        );

        addMonthButton.disabled = true;

        let inflationRate;

        try {
            const response = await fetch(API_URL);

            if (!response.ok) {
                throw new Error("No se pudo consultar la API.");
            }

            const data = await response.json();

            if (!Array.isArray(data)) {
                throw new Error("La respuesta de la API no tiene el formato esperado.");
            }

            const period = formatDateKey(nextMonth);

            const inflationRecord = data.find(item =>
                typeof item.fecha === "string" &&
                item.fecha.slice(0, 7) === period
            );

            if (!inflationRecord || !Number.isFinite(Number(inflationRecord.valor))) {
                throw new Error("No hay un dato válido para ese mes.");
            }

            inflationRate = Number(inflationRecord.valor);

            // Mostrar el porcentaje obtenido en el campo de inflación.
            inflationInput.value = inflationRate.toFixed(2);

        } catch (error) {
            const manualRate = Number(inflationInput.value);

            if (inflationInput.value.trim() === "" || !Number.isFinite(manualRate)) {
                alert(
                    "No se pudo obtener la inflación automáticamente. " +
                    "Puede ingresar el porcentaje manualmente."
                );
                addMonthButton.disabled = false;
                return;
            }

            inflationRate = manualRate;

            alert(
                "No se pudo obtener el dato de la API. " +
                "Se utilizará el porcentaje ingresado manualmente."
            );
        }

        addRow(nextMonth, lastFinalValue, inflationRate);

        addMonthButton.disabled = false;
    });

    function addRow(date, initialValue, inflationRate) {
        const increment = (initialValue * inflationRate) / 100;
        const finalValue = initialValue + increment;

        const row = document.createElement("tr");

        // Guardamos la fecha y el valor final como datos internos.
        row.dataset.month = formatDateKey(date);
        row.dataset.finalValue = finalValue;

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
        if (rentTableBody.rows.length === 0) {
            return;
        }

        rentTableBody.deleteRow(-1);

        if (rentTableBody.rows.length > 0) {
            const lastRow = rentTableBody.lastElementChild;
            const [year, month] = lastRow.dataset.month.split("-").map(Number);

            lastMonth = new Date(year, month - 1, 1);
            lastFinalValue = Number(lastRow.dataset.finalValue);
        } else {
            lastMonth = null;
            lastFinalValue = null;
        }
    });

    exportButton.addEventListener("click", () => {
        const table = document.getElementById("rent-table");
        const ws = XLSX.utils.table_to_sheet(table, { raw: true });

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Cálculo Inflación");

        XLSX.writeFile(wb, "calculo_inflacion.xlsx");
    });
});
