        let file1DataIban = null;
        let file2DataIban = null;
        let file3DataIban = null;
        let file1SelectedColumnIban = null;
        let file2SelectedColumnIban = null;
        let file3SelectedColumnIban = null;      // Колонка з номером ВП у третьому файлі
        let file1VpSelectedColumnIban = null;    // Колонка з номером ВП в основному файлі
        let resultWorkbookIban = null;

        document.getElementById('file1Iban').addEventListener('change', function (e) {
            handleFileUploadIban(e.target.files[0], 1);
        });

        document.getElementById('file2Iban').addEventListener('change', function (e) {
            handleFileUploadIban(e.target.files[0], 2);
        });

        document.getElementById('file3Iban').addEventListener('change', function (e) {
            handleFileUploadIban(e.target.files[0], 3);
        });

        document.getElementById('file1ColumnIban').addEventListener('change', function (e) {
            handleColumnSelectionIban(1, e.target.value);
        });

        document.getElementById('file2ColumnIban').addEventListener('change', function (e) {
            handleColumnSelectionIban(2, e.target.value);
        });

        document.getElementById('file3ColumnIban').addEventListener('change', function (e) {
            handleColumnSelectionIban(3, e.target.value);
        });

        document.getElementById('file1VpColumnIban').addEventListener('change', function (e) {
            file1VpSelectedColumnIban = e.target.value === '' ? null : parseInt(e.target.value);
            checkIfReadyToGenerateIban();
        });

        document.getElementById('generateBtnIban').addEventListener('click', generateResultsIban);
        document.getElementById('downloadBtnIban').addEventListener('click', downloadResultsIban);

        // Заповнює <select> переліком колонок на основі прочитаних рядків
        function populateColumnOptionsIban(select, jsonData, placeholder) {
            const maxCols = Math.max(...jsonData.map(row => row.length));
            select.innerHTML = `<option value="">${placeholder}</option>`;
            for (let i = 0; i < maxCols; i++) {
                const colLetter = String.fromCharCode(65 + i);
                const firstValue = jsonData.length > 1 && jsonData[1][i] ? jsonData[1][i] : '';
                select.innerHTML += `<option value="${i}">${colLetter} - ${firstValue}</option>`;
            }
        }

        function handleFileUploadIban(file, fileNumber) {
            if (!file) return;

            const label = document.getElementById(`label${fileNumber}Iban`);
            const select = document.getElementById(`file${fileNumber}ColumnIban`);

            const reader = new FileReader();

            reader.onload = function (e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                    const fileData = {
                        name: file.name,
                        workbook: workbook,
                        sheetName: firstSheetName,
                        allRows: jsonData
                    };

                    if (fileNumber === 1) {
                        file1DataIban = fileData;
                        file1SelectedColumnIban = null;
                    } else if (fileNumber === 2) {
                        file2DataIban = fileData;
                        file2SelectedColumnIban = null;
                    } else {
                        file3DataIban = fileData;
                        file3SelectedColumnIban = null;
                    }

                    label.classList.add('has-file');
                    const rowCount = jsonData.length - 1;
                    label.innerHTML = `
                        <div class="file-icon">
                            <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7" />
                            </svg>
                        </div>
                        <div class="file-details">
                            <div class="file-name">${file.name}</div>
                            <div class="file-info">${rowCount.toLocaleString('uk-UA')} записів завантажено</div>
                        </div>
                    `;

                    const placeholder = fileNumber === 3
                        ? 'Оберіть колонку з номером ВП (у цьому файлі)'
                        : 'Оберіть колонку з IBAN';
                    populateColumnOptionsIban(select, jsonData, placeholder);
                    select.classList.remove('hidden');

                    // Перший файл також постачає колонки для зіставлення за номером ВП
                    if (fileNumber === 1) {
                        const vpSelect = document.getElementById('file1VpColumnIban');
                        populateColumnOptionsIban(vpSelect, jsonData, 'Оберіть колонку з номером ВП (в основному файлі)');
                        file1VpSelectedColumnIban = null;
                        // Якщо третій файл уже завантажено — показуємо вибір колонки ВП основного файлу
                        if (file3DataIban) {
                            vpSelect.classList.remove('hidden');
                        }
                    }

                    // Третій файл потребує вказати колонку ВП в основному файлі
                    if (fileNumber === 3) {
                        const vpSelect = document.getElementById('file1VpColumnIban');
                        if (file1DataIban) {
                            vpSelect.classList.remove('hidden');
                        }
                    }

                    checkIfReadyToGenerateIban();

                } catch (error) {
                    alert(`Помилка читання файлу: ${error.message}`);
                }
            };

            reader.readAsArrayBuffer(file);
        }

        function handleColumnSelectionIban(fileNumber, columnIndex) {
            const colIndex = columnIndex === '' ? null : parseInt(columnIndex);

            if (fileNumber === 1) {
                file1SelectedColumnIban = colIndex;
            } else if (fileNumber === 2) {
                file2SelectedColumnIban = colIndex;
            } else {
                file3SelectedColumnIban = colIndex;
            }

            checkIfReadyToGenerateIban();
        }

        function checkIfReadyToGenerateIban() {
            const baseReady = file1DataIban && file2DataIban &&
                file1SelectedColumnIban !== null &&
                file2SelectedColumnIban !== null;

            // Третій файл опційний, але якщо завантажений — потрібні обидві колонки ВП
            let file3Ready = true;
            if (file3DataIban) {
                file3Ready = file3SelectedColumnIban !== null && file1VpSelectedColumnIban !== null;
            }

            const canGenerate = baseReady && file3Ready;
            document.getElementById('generateBtnIban').disabled = !canGenerate;

            if (canGenerate) {
                updateStatsIban();
            }
        }

        // Чи задіяне виключення за ВП (третій файл повністю налаштований)
        function isVpExclusionActiveIban() {
            return !!file3DataIban &&
                file3SelectedColumnIban !== null &&
                file1VpSelectedColumnIban !== null;
        }

        // Збирає множину номерів ВП із третього файлу
        function buildVpExclusionSetIban() {
            const set = new Set();
            file3DataIban.allRows.forEach((row, idx) => {
                if (idx > 0 && row[file3SelectedColumnIban]) {
                    set.add(String(row[file3SelectedColumnIban]).trim());
                }
            });
            return set;
        }

        function updateStatsIban() {
            if (!file1DataIban || !file2DataIban ||
                file1SelectedColumnIban === null || file2SelectedColumnIban === null) return;

            const ibanExclusionSet = new Set();
            file2DataIban.allRows.forEach((row, idx) => {
                if (idx > 0 && row[file2SelectedColumnIban]) {
                    ibanExclusionSet.add(String(row[file2SelectedColumnIban]).trim());
                }
            });

            const vpActive = isVpExclusionActiveIban();
            const vpExclusionSet = vpActive ? buildVpExclusionSetIban() : new Set();

            let totalCount = 0;
            let excludedCount = 0;

            file1DataIban.allRows.forEach((row, idx) => {
                if (idx === 0) return;

                const iban = row[file1SelectedColumnIban] ? String(row[file1SelectedColumnIban]).trim() : '';
                if (!iban) return;
                totalCount++;

                const excludedByIban = ibanExclusionSet.has(iban);

                let excludedByVp = false;
                if (vpActive) {
                    const vp = row[file1VpSelectedColumnIban] != null ? String(row[file1VpSelectedColumnIban]).trim() : '';
                    excludedByVp = vp !== '' && vpExclusionSet.has(vp);
                }

                if (excludedByIban || excludedByVp) excludedCount++;
            });

            const toExcludeCount = ibanExclusionSet.size + (vpActive ? vpExclusionSet.size : 0);
            const resultCount = totalCount - excludedCount;

            document.getElementById('statFile1Iban').textContent = totalCount.toLocaleString('uk-UA');
            document.getElementById('statFile2Iban').textContent = toExcludeCount.toLocaleString('uk-UA');
            document.getElementById('statExcludedIban').textContent = excludedCount.toLocaleString('uk-UA');
            document.getElementById('statResultIban').textContent = resultCount.toLocaleString('uk-UA');
        }

        async function generateResultsIban() {
            showProgress('Обробка файлів', 'Фільтрація IBAN номерів', [
                { title: 'Читання файлів', desc: 'Завантаження даних з Excel' },
                { title: 'Аналіз виключень', desc: 'Збір рахунків та ВП' },
                { title: 'Фільтрація', desc: 'Виключення рахунків' },
                { title: 'Створення файлу', desc: 'Генерація результату' }
            ]);

            try {
                updateProgress(10, 1);
                await delay(300);

                updateProgress(25, 1);
                await delay(300);

                updateProgress(30, 2);
                await delay(300);

                const values2Set = new Set();
                file2DataIban.allRows.forEach((row, idx) => {
                    if (idx > 0 && row[file2SelectedColumnIban]) {
                        values2Set.add(String(row[file2SelectedColumnIban]).trim());
                    }
                });

                const vpActive = isVpExclusionActiveIban();
                const vpExclusionSet = vpActive ? buildVpExclusionSetIban() : new Set();

                updateProgress(50, 2);
                await delay(400);

                updateProgress(55, 3);
                await delay(300);

                const resultRows = [file1DataIban.allRows[0]];
                file1DataIban.allRows.forEach((row, idx) => {
                    if (idx > 0) {
                        const value = row[file1SelectedColumnIban] ? String(row[file1SelectedColumnIban]).trim() : '';
                        if (!value) return;

                        const excludedByIban = values2Set.has(value);

                        let excludedByVp = false;
                        if (vpActive) {
                            const vp = row[file1VpSelectedColumnIban] != null ? String(row[file1VpSelectedColumnIban]).trim() : '';
                            excludedByVp = vp !== '' && vpExclusionSet.has(vp);
                        }

                        if (!excludedByIban && !excludedByVp) {
                            resultRows.push(row);
                        }
                    }
                });

                updateProgress(75, 3);
                await delay(400);

                updateProgress(80, 4);
                await delay(300);

                const wb = XLSX.utils.book_new();

                const ws1 = XLSX.utils.aoa_to_sheet(file1DataIban.allRows);
                formatIBANColumn(ws1, file1SelectedColumnIban, file1DataIban.allRows.length);
                XLSX.utils.book_append_sheet(wb, ws1, 'Всі рахунки');

                updateProgress(85, 4);
                await delay(200);

                const ws2 = XLSX.utils.aoa_to_sheet(file2DataIban.allRows);
                formatIBANColumn(ws2, file2SelectedColumnIban, file2DataIban.allRows.length);
                XLSX.utils.book_append_sheet(wb, ws2, 'Зарплатні');

                // Аркуш зі списком ВП (лише якщо третій файл задіяно)
                if (vpActive) {
                    const wsVp = XLSX.utils.aoa_to_sheet(file3DataIban.allRows);
                    formatIBANColumn(wsVp, file3SelectedColumnIban, file3DataIban.allRows.length);
                    XLSX.utils.book_append_sheet(wb, wsVp, 'Виключення ВП');
                }

                updateProgress(90, 4);
                await delay(200);

                const resultRowsRenamed = resultRows.map((row, idx) => {
                    if (idx === 0) {
                        const newRow = [...row];
                        if (newRow.length > 0) newRow[0] = 'asvp_id';
                        if (newRow.length > 1) newRow[1] = 'iban';
                        return newRow;
                    }
                    return row;
                });

                const ws3 = XLSX.utils.aoa_to_sheet(resultRowsRenamed);
                formatIBANColumn(ws3, file1SelectedColumnIban, resultRowsRenamed.length);
                XLSX.utils.book_append_sheet(wb, ws3, 'Результат');

                resultWorkbookIban = wb;

                updateProgress(100, 4);
                await delay(500);

                hideProgress();

                document.getElementById('successMessageIban').classList.add('show');
                document.getElementById('downloadBtnIban').classList.remove('hidden');

            } catch (error) {
                hideProgress();
                alert(`Помилка: ${error.message}`);
            }
        }
        function downloadResultsIban() {
            if (!resultWorkbookIban) return;

            try {
                const wbout = XLSX.write(resultWorkbookIban, {
                    bookType: 'xlsx',
                    type: 'array',
                    cellDates: false,
                    bookSST: false
                });

                const blob = new Blob([wbout], { type: 'application/octet-stream' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'IBAN_результат.xlsx';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);

            } catch (error) {
                alert(`Помилка завантаження: ${error.message}`);
            }
        }
