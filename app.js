/* ==========================================
   Web Application Logic for Nutrition Report (3-Day Version)
   ========================================== */

document.addEventListener('DOMContentLoaded', () => {
    // UI Elements
    const uploadSection = document.getElementById('upload-section');
    const reportSection = document.getElementById('report-section');
    
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const parseLoading = document.getElementById('parse-loading');
    const uploadError = document.getElementById('upload-error');
    
    const backBtn = document.getElementById('back-btn');
    const pdfSummaryBtn = document.getElementById('pdf-summary-btn');
    const pdfSingleBtn = document.getElementById('pdf-single-btn');
    const pdfAllBtn = document.getElementById('pdf-all-btn');
    
    const patientNameInput = document.getElementById('patient-name');
    const summaryPatientName = document.getElementById('summary-patient-name');
    const reportDateDisplay = document.getElementById('report-date-display');
    const adviceTextarea = document.getElementById('advice-textarea');
    const summaryAdviceTextarea = document.getElementById('summary-advice-textarea');
    const summaryAdvicePrintDisplay = document.getElementById('summary-advice-print-display');
    const dateTabsContainer = document.getElementById('date-tabs-container');
    const dayNumberInput = document.getElementById('day-number-input');

    // Synchronize patient name between summary report and single-day report
    if (patientNameInput && summaryPatientName) {
        patientNameInput.addEventListener('input', (e) => {
            summaryPatientName.value = e.target.value;
        });
        summaryPatientName.addEventListener('input', (e) => {
            patientNameInput.value = e.target.value;
        });
    }

    // Synchronize day number input to global state
    dayNumberInput.addEventListener('input', (e) => {
        if (activeDateKey && parsedData && parsedData.days[activeDateKey]) {
            parsedData.days[activeDateKey].dayLabel = e.target.value;
        }
    });

    // Global state
    let parsedData = null; // Stores all days: { days: { "2026-06-22": { meals, totals } }, dates: [...], summaryAdvice: "" }
    let activeTab = 'summary'; // 'summary' or dateKey
    let activeDateKey = null; // Currently selected date key for single report
    let pfcChartInstance = null; // Single chart instance for preview screen
    let summaryChartInstance = null; // Chart instance for 3-day summary report

    // Synchronize advice textarea input to preview div
    adviceTextarea.addEventListener('input', (e) => {
        const display = document.getElementById('advice-print-display');
        if (display) {
            display.textContent = e.target.value || "アドバイスは未記入です。";
        }
    });

    // Synchronize summary advice textarea input to preview div and state
    if (summaryAdviceTextarea) {
        summaryAdviceTextarea.addEventListener('input', (e) => {
            if (parsedData) {
                parsedData.summaryAdvice = e.target.value;
            }
            if (summaryAdvicePrintDisplay) {
                summaryAdvicePrintDisplay.textContent = e.target.value || "アドバイスは未記入です。";
            }
        });
    }



    // Initialize Lucide Icons
    lucide.createIcons();

    // ==========================================
    // 1. EXCEL UPLOAD & ZIP PARSING
    // ==========================================
    // Drag & Drop handlers
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
        }, false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleExcelFile(files[0]);
        }
    });

    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleExcelFile(e.target.files[0]);
        }
    });

    // Excel analysis core
    async function handleExcelFile(file) {
        if (!file.name.endsWith('.xlsx')) {
            showUploadError('Excelファイル（.xlsx）をアップロードしてください。');
            return;
        }

        showLoading(true);
        showUploadError('');

        try {
            // 1. Extract images using JSZip
            const zip = await JSZip.loadAsync(file);
            const mediaFiles = {};
            
            // Extract media binaries
            for (let key in zip.files) {
                if (key.startsWith('xl/media/')) {
                    const base64 = await zip.files[key].async('base64');
                    const relName = key.replace('xl/media/', '../media/');
                    mediaFiles[relName] = `data:image/jpeg;base64,${base64}`;
                }
            }

            // Parse drawings to match row indices to extracted image data URLs
            const rowToImageMap = {};
            const relsFile = zip.file('xl/drawings/_rels/drawing1.xml.rels');
            const drawingFile = zip.file('xl/drawings/drawing1.xml');

            if (relsFile && drawingFile) {
                const parser = new DOMParser();
                const relsText = await relsFile.async('string');
                const relsDoc = parser.parseFromString(relsText, 'application/xml');
                const relationshipEls = relsDoc.getElementsByTagName('Relationship');
                
                const ridMap = {};
                for (let i = 0; i < relationshipEls.length; i++) {
                    const rel = relationshipEls[i];
                    ridMap[rel.getAttribute('Id')] = rel.getAttribute('Target');
                }

                const drawingText = await drawingFile.async('string');
                const drawingDoc = parser.parseFromString(drawingText, 'application/xml');
                const anchors = drawingDoc.querySelectorAll('*|oneCellAnchor');

                anchors.forEach(anchor => {
                    const fromEl = anchor.querySelector('*|from');
                    if (fromEl) {
                        const colEl = fromEl.querySelector('*|col');
                        const rowEl = fromEl.querySelector('*|row');
                        
                        if (colEl && rowEl) {
                            const colVal = parseInt(colEl.textContent, 10);
                            const rowVal = parseInt(rowEl.textContent, 10);
                            
                            const picEl = anchor.querySelector('*|pic');
                            if (picEl) {
                                const blipEl = picEl.querySelector('*|blip');
                                if (blipEl) {
                                    let embedRid = '';
                                    for (let attr of blipEl.attributes) {
                                        if (attr.name.endsWith('embed')) {
                                            embedRid = attr.value;
                                            break;
                                        }
                                    }
                                    
                                    const targetImage = ridMap[embedRid];
                                    const imageData = mediaFiles[targetImage];
                                    
                                    if (imageData) {
                                        // Fix 2-row layout offset: Excel Row 4 (rowVal 3) -> Data index 0
                                        const targetDataIndex = rowVal - 3;
                                        if (targetDataIndex >= 0) {
                                            rowToImageMap[targetDataIndex] = imageData;
                                        }
                                    }
                                }
                            }
                        }
                    }
                });
            }

            // 2. Parse sheet data using SheetJS
            const arrayBuffer = await readFileAsArrayBuffer(file);
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            const rawRows = XLSX.utils.sheet_to_json(worksheet);
            if (rawRows.length === 0) {
                throw new Error('Excelシート内にデータが見つかりません。');
            }

            // Organize data by Date
            const daysData = {};

            rawRows.forEach((row, index) => {
                if (!row['日時']) return;
                
                const d = new Date(row['日時']);
                if (isNaN(d.getTime())) return;

                // Date Key format: YYYY-MM-DD
                const dateKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                const hours = d.getHours();

                // Initialize date group if not exists
                if (!daysData[dateKey]) {
                    daysData[dateKey] = {
                        displayDate: `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`,
                        meals: {
                            breakfast: { items: [], kcal: 0, protein: 0, fat: 0, carb: 0, salt: 0, images: [] },
                            lunch: { items: [], kcal: 0, protein: 0, fat: 0, carb: 0, salt: 0, images: [] },
                            dinner: { items: [], kcal: 0, protein: 0, fat: 0, carb: 0, salt: 0, images: [] },
                            other: { items: [], kcal: 0, protein: 0, fat: 0, carb: 0, salt: 0, images: [] }
                        },
                        totals: {
                            kcal: 0, protein: 0, fat: 0, carb: 0, fiber: 0,
                            calcium: 0, iron: 0, salt: 0,
                            vitA: 0, vitD: 0, vitB1: 0, vitB2: 0, vitB6: 0, vitC: 0
                        },
                        advice: "" // Holds personalized advice per day
                    };
                }

                // Determine meal type by time
                let mealType = 'other';
                if (hours >= 5 && hours < 11) {
                    mealType = 'breakfast';
                } else if (hours >= 11 && hours < 16) {
                    mealType = 'lunch';
                } else if (hours >= 18 && hours < 24) {
                    mealType = 'dinner';
                } else {
                    mealType = 'other';
                }

                // Add picture with DUPLICATE FILTERING (Ensure unique images per meal)
                const imageSrc = rowToImageMap[index] || null;
                if (imageSrc) {
                    if (!daysData[dateKey].meals[mealType].images.includes(imageSrc)) {
                        daysData[dateKey].meals[mealType].images.push(imageSrc);
                    }
                }

                // Food details
                const foodName = row['料理名'] || '無題';
                const memo = row['メモ'] || '';
                const unit = row['単位'] || '';
                
                daysData[dateKey].meals[mealType].items.push({
                    name: foodName,
                    memo: memo,
                    unit: unit
                });

                // Nutritional parsing
                const rowKcal = parseFloat(row['エネルギー（kcal）(kcal)']) || 0;
                const rowProtein = parseFloat(row['たんぱく質(g)']) || 0;
                const rowFat = parseFloat(row['脂質(g)']) || 0;
                const rowCarb = parseFloat(row['炭水化物(g)']) || 0;
                const rowFiber = parseFloat(row['食物繊維総量(g)']) || 0;
                const rowCalcium = parseFloat(row['カルシウム(mg)']) || 0;
                const rowIron = parseFloat(row['鉄(mg)']) || 0;
                const rowSalt = parseFloat(row['食塩相当量(g)']) || 0;

                const rowVitA = parseFloat(row['ビタミンA(レチノール当量)(µg)']) || 0;
                const rowVitD = parseFloat(row['ビタミンD(µg)']) || 0;
                const rowVitB1 = parseFloat(row['ビタミンB1(mg)']) || 0;
                const rowVitB2 = parseFloat(row['ビタミンB2(mg)']) || 0;
                const rowVitB6 = parseFloat(row['ビタミンB6(mg)']) || 0;
                const rowVitC = parseFloat(row['ビタミンC(mg)']) || 0;

                // Add to section totals (Kcal, P, F, C, Salt now categorized by meals!)
                daysData[dateKey].meals[mealType].kcal += rowKcal;
                daysData[dateKey].meals[mealType].protein += rowProtein;
                daysData[dateKey].meals[mealType].fat += rowFat;
                daysData[dateKey].meals[mealType].carb += rowCarb;
                daysData[dateKey].meals[mealType].salt += rowSalt;

                // Add to daily totals
                const t = daysData[dateKey].totals;
                t.kcal += rowKcal;
                t.protein += rowProtein;
                t.fat += rowFat;
                t.carb += rowCarb;
                t.fiber += rowFiber;
                t.calcium += rowCalcium;
                t.iron += rowIron;
                t.salt += rowSalt;

                t.vitA += rowVitA;
                t.vitD += rowVitD;
                t.vitB1 += rowVitB1;
                t.vitB2 += rowVitB2;
                t.vitB6 += rowVitB6;
                t.vitC += rowVitC;
            });

            const sortedDates = Object.keys(daysData).sort();
            if (sortedDates.length === 0) {
                throw new Error('日付を特定できるデータが見つかりませんでした。');
            }

            parsedData = {
                days: daysData,
                dates: sortedDates,
                summaryAdvice: ""
            };

            fileInput.value = '';
            showLoading(false);
            
            // Build date tabs and render 3-day summary report first (User default: 1-page 3-day summary)
            buildDateTabs();
            selectSummaryTab();
            showScreen('report-section');

        } catch (err) {
            console.error(err);
            showLoading(false);
            showUploadError(`エラーが発生しました: ${err.message || 'ファイルのパースに失敗しました。'}`);
        }
    }

    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(new Error('ファイルの読み込みに失敗しました。'));
            reader.readAsArrayBuffer(file);
        });
    }

    // ==========================================
    // 2. TABS MANAGEMENT
    // ==========================================
    function saveCurrentState() {
        if (activeTab === 'summary') {
            if (summaryAdviceTextarea && parsedData) {
                parsedData.summaryAdvice = summaryAdviceTextarea.value;
            }
        } else if (activeDateKey && parsedData && parsedData.days[activeDateKey]) {
            if (adviceTextarea) {
                parsedData.days[activeDateKey].advice = adviceTextarea.value;
            }
        }
    }

    function buildDateTabs() {
        dateTabsContainer.innerHTML = '';

        // 1. "3日間まとめ (1枚)" Tab
        const summaryBtn = document.createElement('button');
        summaryBtn.className = 'tab-btn active';
        summaryBtn.dataset.tab = 'summary';
        summaryBtn.innerHTML = '<i data-lucide="sparkles" style="width:14px;height:14px;vertical-align:-2px;margin-right:4px;"></i>3日間まとめ (1枚)';
        summaryBtn.addEventListener('click', () => {
            saveCurrentState();
            selectSummaryTab();
        });
        dateTabsContainer.appendChild(summaryBtn);

        // 2. Day-by-day tabs
        parsedData.dates.forEach((dateKey, index) => {
            const btn = document.createElement('button');
            btn.className = 'tab-btn';
            btn.dataset.tab = dateKey;
            
            const d = new Date(dateKey);
            btn.textContent = `${index + 1}日目 (${d.getMonth() + 1}月${d.getDate()}日)`;
            
            btn.addEventListener('click', () => {
                saveCurrentState();
                selectDay(dateKey);
            });
            dateTabsContainer.appendChild(btn);
        });

        lucide.createIcons();
    }

    function selectSummaryTab() {
        activeTab = 'summary';
        
        // Active tab styling
        document.querySelectorAll('.tab-btn').forEach(btn => {
            if (btn.dataset.tab === 'summary') {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });

        // Switch visible containers
        const summaryReport = document.getElementById('three-days-summary-report');
        const singleReport = document.getElementById('printable-report');
        if (summaryReport) {
            summaryReport.style.display = 'block';
            summaryReport.classList.add('active');
        }
        if (singleReport) {
            singleReport.style.display = 'none';
            singleReport.classList.remove('active');
        }

        // Toggle action buttons
        if (pdfSummaryBtn) pdfSummaryBtn.style.display = 'inline-flex';
        if (pdfSingleBtn) pdfSingleBtn.style.display = 'none';
        if (pdfAllBtn) pdfAllBtn.style.display = 'none';

        render3DaysSummaryReport();
    }

    function selectDay(dateKey) {
        activeTab = dateKey;
        activeDateKey = dateKey;
        
        // Active tab styling
        document.querySelectorAll('.tab-btn').forEach(btn => {
            if (btn.dataset.tab === dateKey) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });

        // Switch visible containers
        const summaryReport = document.getElementById('three-days-summary-report');
        const singleReport = document.getElementById('printable-report');
        if (summaryReport) {
            summaryReport.style.display = 'none';
            summaryReport.classList.remove('active');
        }
        if (singleReport) {
            singleReport.style.display = 'block';
            singleReport.classList.add('active');
        }

        // Toggle action buttons
        if (pdfSummaryBtn) pdfSummaryBtn.style.display = 'none';
        if (pdfSingleBtn) pdfSingleBtn.style.display = 'inline-flex';
        if (pdfAllBtn) pdfAllBtn.style.display = 'inline-flex';

        // Set day number input (e.g. 2 → "2日目")
        const dayIndex = parsedData.dates.indexOf(dateKey) + 1;
        if (dayNumberInput) dayNumberInput.value = String(dayIndex);

        renderReportForDate(dateKey);
    }

    // ==========================================
    // 2.5 3-DAY INTEGRATED SUMMARY REPORT RENDERING
    // ==========================================
    function render3DaysSummaryReport() {
        if (!parsedData || !parsedData.dates || parsedData.dates.length === 0) return;

        // 1. Period & Patient metadata
        const firstDate = new Date(parsedData.dates[0]);
        const lastDate = new Date(parsedData.dates[parsedData.dates.length - 1]);
        const periodDisplay = document.getElementById('summary-period-display');
        if (periodDisplay) {
            if (parsedData.dates.length === 1) {
                periodDisplay.textContent = `${firstDate.getFullYear()}年${firstDate.getMonth() + 1}月${firstDate.getDate()}日`;
            } else {
                periodDisplay.textContent = `${firstDate.getFullYear()}年${firstDate.getMonth() + 1}月${firstDate.getDate()}日 〜 ${lastDate.getMonth() + 1}月${lastDate.getDate()}日 (${parsedData.dates.length}日間)`;
            }
        }

        // Sync patient name
        if (patientNameInput && summaryPatientName) {
            summaryPatientName.value = patientNameInput.value;
        }

        // 2. Meals Grid (3 Days horizontal columns)
        const summaryMealsGrid = document.getElementById('summary-meals-grid');
        summaryMealsGrid.innerHTML = '';

        const mealTypes = [
            { type: 'breakfast', name: '朝食', icon: '☀️', bgClass: 'bg-breakfast' },
            { type: 'lunch', name: '昼食', icon: '🕛', bgClass: 'bg-lunch' },
            { type: 'dinner', name: '夕食', icon: '🌙', bgClass: 'bg-dinner' },
            { type: 'other', name: 'その他', icon: '☕', bgClass: 'bg-other' }
        ];

        parsedData.dates.forEach((dateKey, index) => {
            const dayData = parsedData.days[dateKey];
            const d = new Date(dateKey);
            const dayCol = document.createElement('div');
            dayCol.className = 'summary-day-col';

            // Column Header
            const dayHeader = document.createElement('div');
            dayHeader.className = 'summary-day-header';
            dayHeader.innerHTML = `
                <span>${index + 1}日目 <span style="font-weight:400;font-size:10.5px;">(${d.getMonth() + 1}月${d.getDate()}日)</span></span>
                <span class="day-sub-kcal">${Math.round(dayData.totals.kcal)} kcal</span>
            `;
            dayCol.appendChild(dayHeader);

            // Meal blocks
            mealTypes.forEach(m => {
                const mealData = dayData.meals[m.type];
                const mealBlock = document.createElement('div');
                mealBlock.className = 'summary-meal-block';

                // Title bar
                const mealTitle = document.createElement('div');
                mealTitle.className = `summary-meal-title ${m.bgClass}`;
                mealTitle.innerHTML = `
                    <span>${m.icon} ${m.name}</span>
                    <span class="meal-sub-val">${Math.round(mealData.kcal)} kcal</span>
                `;
                mealBlock.appendChild(mealTitle);

                // Body (Photo thumbnail + Compact food items)
                const mealBody = document.createElement('div');
                mealBody.className = 'summary-meal-body';

                // Thumbnail box
                const imgBox = document.createElement('div');
                imgBox.className = 'summary-meal-images';
                if (mealData.images && mealData.images.length > 0) {
                    const img = document.createElement('img');
                    img.src = mealData.images[0];
                    img.alt = `${m.name}写真`;
                    imgBox.appendChild(img);
                } else {
                    imgBox.innerHTML = '<span class="summary-no-img">写真なし</span>';
                }
                mealBody.appendChild(imgBox);

                // Food list (Extremely compact, food name + unit, no wasted width)
                const foodList = document.createElement('div');
                foodList.className = 'summary-food-list';
                if (!mealData.items || mealData.items.length === 0) {
                    foodList.innerHTML = '<div class="summary-food-empty">記録なし</div>';
                } else {
                    mealData.items.forEach(item => {
                        const foodItem = document.createElement('div');
                        foodItem.className = 'summary-food-item';
                        
                        const nameSpan = document.createElement('span');
                        nameSpan.className = 'summary-food-name';
                        nameSpan.textContent = item.name;
                        nameSpan.title = item.name;

                        foodItem.appendChild(nameSpan);

                        if (item.unit) {
                            const unitSpan = document.createElement('span');
                            unitSpan.className = 'summary-food-unit';
                            unitSpan.textContent = `(${item.unit})`;
                            foodItem.appendChild(unitSpan);
                        }
                        foodList.appendChild(foodItem);
                    });
                }
                mealBody.appendChild(foodList);

                mealBlock.appendChild(mealBody);
                dayCol.appendChild(mealBlock);
            });

            summaryMealsGrid.appendChild(dayCol);
        });

        // 3. Nutrients Comparison Table
        const tbody = document.getElementById('summary-nutrients-tbody');
        tbody.innerHTML = '';

        const nutrientDefinitions = [
            { key: 'kcal', name: 'エネルギー', unit: 'kcal', digits: 0, ref: '1800〜2200' },
            { key: 'protein', name: 'たんぱく質', unit: 'g', digits: 1, ref: '50〜65g' },
            { key: 'fat', name: '脂質', unit: 'g', digits: 1, ref: '20〜30%' },
            { key: 'carb', name: '炭水化物', unit: 'g', digits: 1, ref: '50〜65%' },
            { key: 'fiber', name: '食物繊維', unit: 'g', digits: 1, ref: '18g以上' },
            { key: 'calcium', name: 'カルシウム', unit: 'mg', digits: 0, ref: '650〜800' },
            { key: 'iron', name: '鉄', unit: 'mg', digits: 1, ref: '6.5〜10.5' },
            { key: 'salt', name: '食塩相当量', unit: 'g', digits: 1, ref: '7.5g未満' }
        ];

        const numDays = parsedData.dates.length;

        // Dynamic TH headers
        parsedData.dates.forEach((dateKey, idx) => {
            const thEl = document.getElementById(`th-day-${idx + 1}`);
            if (thEl) {
                const d = new Date(dateKey);
                thEl.textContent = `${idx + 1}日目(${d.getMonth() + 1}/${d.getDate()})`;
            }
        });

        nutrientDefinitions.forEach(def => {
            const tr = document.createElement('tr');
            let rowHtml = `<td>${def.name} <span style="font-weight:400;font-size:8px;color:#718096;">(${def.unit})</span></td>`;

            let sum = 0;
            parsedData.dates.forEach(dateKey => {
                const val = parsedData.days[dateKey].totals[def.key] || 0;
                sum += val;
                const displayVal = def.digits === 0 ? Math.round(val) : val.toFixed(def.digits);
                rowHtml += `<td>${displayVal}</td>`;
            });

            // 3-day average
            const avg = numDays > 0 ? sum / numDays : 0;
            const displayAvg = def.digits === 0 ? Math.round(avg) : avg.toFixed(def.digits);
            rowHtml += `<td class="col-avg">${displayAvg}</td>`;
            rowHtml += `<td class="col-ref">${def.ref}</td>`;

            tr.innerHTML = rowHtml;
            tbody.appendChild(tr);
        });

        // 4. Vitamins 3-day Average Summary Bar
        const vitBar = document.getElementById('summary-vitamins-bar');
        if (vitBar) {
            const vits = [
                { key: 'vitA', label: 'ビタミンA', unit: 'µg', digits: 0 },
                { key: 'vitD', label: 'ビタミンD', unit: 'µg', digits: 1 },
                { key: 'vitB1', label: 'ビタミンB1', unit: 'mg', digits: 2 },
                { key: 'vitB2', label: 'ビタミンB2', unit: 'mg', digits: 2 },
                { key: 'vitB6', label: 'ビタミンB6', unit: 'mg', digits: 2 },
                { key: 'vitC', label: 'ビタミンC', unit: 'mg', digits: 0 }
            ];

            vitBar.innerHTML = vits.map(v => {
                let sum = 0;
                parsedData.dates.forEach(dateKey => {
                    sum += (parsedData.days[dateKey].totals[v.key] || 0);
                });
                const avg = numDays > 0 ? sum / numDays : 0;
                const valStr = v.digits === 0 ? Math.round(avg) : avg.toFixed(v.digits);
                return `
                    <div class="summary-vit-item">
                        <span class="summary-vit-lbl">${v.label} (平均)</span>
                        <span class="summary-vit-val">${valStr} ${v.unit}</span>
                    </div>
                `;
            }).join('');
        }

        // 5. PFC Balance 3-day Average Donut Chart
        let totalP = 0, totalF = 0, totalC = 0;
        parsedData.dates.forEach(dateKey => {
            const t = parsedData.days[dateKey].totals;
            totalP += t.protein;
            totalF += t.fat;
            totalC += t.carb;
        });

        const avgP = numDays > 0 ? totalP / numDays : 0;
        const avgF = numDays > 0 ? totalF / numDays : 0;
        const avgC = numDays > 0 ? totalC / numDays : 0;

        const pKcal = avgP * 4;
        const fKcal = avgF * 9;
        const cKcal = avgC * 4;
        const totalPfcKcal = pKcal + fKcal + cKcal;

        let pPct = 0, fPct = 0, cPct = 0;
        if (totalPfcKcal > 0) {
            pPct = (pKcal / totalPfcKcal) * 100;
            fPct = (fKcal / totalPfcKcal) * 100;
            cPct = (cKcal / totalPfcKcal) * 100;
        }

        const pPctEl = document.getElementById('summary-pfc-p-pct');
        const pValEl = document.getElementById('summary-pfc-p-val');
        const fPctEl = document.getElementById('summary-pfc-f-pct');
        const fValEl = document.getElementById('summary-pfc-f-val');
        const cPctEl = document.getElementById('summary-pfc-c-pct');
        const cValEl = document.getElementById('summary-pfc-c-val');

        if (pPctEl) pPctEl.textContent = `${pPct.toFixed(1)}%`;
        if (pValEl) pValEl.textContent = `${avgP.toFixed(1)}g`;
        if (fPctEl) fPctEl.textContent = `${fPct.toFixed(1)}%`;
        if (fValEl) fValEl.textContent = `${avgF.toFixed(1)}g`;
        if (cPctEl) cPctEl.textContent = `${cPct.toFixed(1)}%`;
        if (cValEl) cValEl.textContent = `${avgC.toFixed(1)}g`;

        const chartCanvas = document.getElementById('summaryPfcChart');
        if (chartCanvas) {
            const ctx = chartCanvas.getContext('2d');
            if (summaryChartInstance) {
                summaryChartInstance.destroy();
            }

            summaryChartInstance = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: ['P (たんぱく質)', 'F (脂質)', 'C (炭水化物)'],
                    datasets: [{
                        data: [pKcal, fKcal, cKcal],
                        backgroundColor: ['#ff6b6b', '#feca57', '#48dbfb'],
                        borderWidth: 1.5,
                        borderColor: '#ffffff'
                    }]
                },
                options: {
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const pct = ((val / totalPfcKcal) * 100).toFixed(1);
                                    return `${context.label}: ${Math.round(val)} kcal (${pct}%)`;
                                }
                            }
                        }
                    },
                    cutout: '65%',
                    responsive: true,
                    maintainAspectRatio: false
                }
            });
        }

        // 6. Advice sync
        if (summaryAdviceTextarea) {
            summaryAdviceTextarea.value = parsedData.summaryAdvice || "";
            if (summaryAdvicePrintDisplay) {
                summaryAdvicePrintDisplay.textContent = parsedData.summaryAdvice || "アドバイスは未記入です。";
            }
        }
    }

    // ==========================================
    // 3. REPORT RENDERING FOR SELECTED DATE
    // ==========================================
    function renderReportForDate(dateKey) {
        const dayData = parsedData.days[dateKey];
        if (!dayData) return;

        // Set metadata
        reportDateDisplay.textContent = dayData.displayDate;
        adviceTextarea.value = dayData.advice || "";
        const advicePrintDisplay = document.getElementById('advice-print-display');
        if (advicePrintDisplay) {
            advicePrintDisplay.textContent = dayData.advice || "アドバイスは未記入です。";
        }

        const dayIndex = parsedData.dates.indexOf(dateKey) + 1;
        // Load or initialize dayLabel in parsedData state
        if (dayData.dayLabel === undefined) {
            dayData.dayLabel = String(dayIndex);
        }
        if (dayNumberInput) {
            dayNumberInput.value = dayData.dayLabel;
        }



        const meals = dayData.meals;
        const totals = dayData.totals;

        // Render meal sections (Food list, Images, and sub-nutrients)
        const mealTypes = ['breakfast', 'lunch', 'dinner', 'other'];
        
        mealTypes.forEach(type => {
            const container = document.getElementById(`list-${type}`);
            const imgContainer = document.getElementById(`images-${type}`);
            const data = meals[type];

            // 1. Food items
            container.innerHTML = '';
            if (data.items.length === 0) {
                container.innerHTML = `<tr><td colspan="3" class="text-muted text-center" style="font-style:italic; padding:12px;">食事データなし</td></tr>`;
            } else {
                data.items.forEach(item => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td style="font-weight: 500;">${item.name}</td>
                        <td class="text-muted">${item.memo}</td>
                        <td style="text-align: right; font-weight: 500;">${item.unit}</td>
                    `;
                    container.appendChild(tr);
                });
            }

            // 2. Images (Unique duplicates filtered already in parse)
            imgContainer.innerHTML = '';
            if (data.images.length === 0) {
                imgContainer.innerHTML = `<span class="no-image-placeholder">写真なし</span>`;
            } else {
                data.images.forEach(imgSrc => {
                    const wrapper = document.createElement('div');
                    wrapper.className = 'meal-img-wrapper';
                    const img = document.createElement('img');
                    img.src = imgSrc;
                    img.className = 'meal-img';
                    img.alt = `${type} image`;
                    wrapper.appendChild(img);
                    imgContainer.appendChild(wrapper);
                });
            }

            // 3. Section sub-totals in nutrition table (Kcal, P, F, C, Salt now dynamic per meal!)
            document.getElementById(`val-${type}-kcal`).textContent = `${Math.round(data.kcal)} kcal`;
            document.getElementById(`val-${type}-protein`).textContent = `${data.protein.toFixed(1)} g`;
            document.getElementById(`val-${type}-fat`).textContent = `${data.fat.toFixed(1)} g`;
            document.getElementById(`val-${type}-carb`).textContent = `${data.carb.toFixed(1)} g`;
            document.getElementById(`val-${type}-salt`).textContent = `${data.salt.toFixed(1)} g`;
        });

        // Render daily totals table values
        document.getElementById('val-total-kcal').textContent = `${Math.round(totals.kcal)} kcal`;
        document.getElementById('val-total-protein').textContent = `${totals.protein.toFixed(1)} g`;
        document.getElementById('val-total-fat').textContent = `${totals.fat.toFixed(1)} g`;
        document.getElementById('val-total-carb').textContent = `${totals.carb.toFixed(1)} g`;
        document.getElementById('val-total-fiber').textContent = `${totals.fiber.toFixed(1)} g`;
        document.getElementById('val-total-calcium').textContent = `${Math.round(totals.calcium)} mg`;
        document.getElementById('val-total-iron').textContent = `${totals.iron.toFixed(1)} mg`;
        document.getElementById('val-total-salt').textContent = `${totals.salt.toFixed(1)} g`;

        // Render extra vitamins
        document.getElementById('val-total-vit-a').textContent = `${Math.round(totals.vitA)}`;
        document.getElementById('val-total-vit-d').textContent = `${totals.vitD.toFixed(1)}`;
        document.getElementById('val-total-vit-b1').textContent = `${totals.vitB1.toFixed(2)}`;
        document.getElementById('val-total-vit-b2').textContent = `${totals.vitB2.toFixed(2)}`;
        document.getElementById('val-total-vit-b6').textContent = `${totals.vitB6.toFixed(2)}`;
        document.getElementById('val-total-vit-c').textContent = `${Math.round(totals.vitC)}`;

        // Render PFC Balance Graph
        renderPfcBalanceGraph(totals.protein, totals.fat, totals.carb);
    }

    function renderPfcBalanceGraph(protein, fat, carb) {
        const pKcal = protein * 4;
        const fKcal = fat * 9;
        const cKcal = carb * 4;
        const totalPfcKcal = pKcal + fKcal + cKcal;

        let pPct = 0, fPct = 0, cPct = 0;
        if (totalPfcKcal > 0) {
            pPct = (pKcal / totalPfcKcal) * 100;
            fPct = (fKcal / totalPfcKcal) * 100;
            cPct = (cKcal / totalPfcKcal) * 100;
        }

        // Set legend details
        document.getElementById('pfc-p-val').textContent = `${protein.toFixed(1)}g (${Math.round(pKcal)} kcal)`;
        document.getElementById('pfc-p-pct').textContent = `${pPct.toFixed(1)}%`;
        
        document.getElementById('pfc-f-val').textContent = `${fat.toFixed(1)}g (${Math.round(fKcal)} kcal)`;
        document.getElementById('pfc-f-pct').textContent = `${fPct.toFixed(1)}%`;
        
        document.getElementById('pfc-c-val').textContent = `${carb.toFixed(1)}g (${Math.round(cKcal)} kcal)`;
        document.getElementById('pfc-c-pct').textContent = `${cPct.toFixed(1)}%`;

        // Draw Chart.js doughnut chart
        const ctx = document.getElementById('pfcChart').getContext('2d');
        
        if (pfcChartInstance) {
            pfcChartInstance.destroy();
        }

        pfcChartInstance = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['P (たんぱく質)', 'F (脂質)', 'C (炭水化物)'],
                datasets: [{
                    data: [pKcal, fKcal, cKcal],
                    backgroundColor: ['#ff6b6b', '#feca57', '#48dbfb'],
                    borderWidth: 2,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const val = context.raw;
                                const pct = ((val / totalPfcKcal) * 100).toFixed(1);
                                return `${context.label}: ${Math.round(val)} kcal (${pct}%)`;
                            }
                        }
                    }
                },
                cutout: '65%',
                responsive: true,
                maintainAspectRatio: false
            }
        });
    }

    // ==========================================
    // 4. PDF GENERATION (SUMMARY, SINGLE DAY & ALL DAYS)
    // ==========================================
    // 4.0 Download 3-Day Integrated Summary (Exact 1-page A4 Landscape)
    if (pdfSummaryBtn) {
        pdfSummaryBtn.addEventListener('click', () => {
            // Save advice first
            if (summaryAdviceTextarea && parsedData) {
                parsedData.summaryAdvice = summaryAdviceTextarea.value;
            }

            const element = document.getElementById('three-days-summary-report');
            const patientName = (summaryPatientName ? summaryPatientName.value.trim() : '') || (patientNameInput ? patientNameInput.value.trim() : '');
            
            const firstDate = parsedData && parsedData.dates ? new Date(parsedData.dates[0]) : null;
            const dateStr = firstDate ? `${firstDate.getMonth() + 1}月` : '食事記録';
            const filename = `栄養食事指導参考資料_${patientName ? patientName + '様' : '食事記録'}_3日間まとめ.pdf`;

            const opt = {
                margin: [4, 6, 4, 6], // Top/Bottom 4mm, Left/Right 6mm
                filename: filename,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true, logging: false, letterRendering: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' },
                pagebreak: { mode: ['css', 'legacy'] }
            };

            // Switch textarea to preview div during PDF compilation
            if (summaryAdvicePrintDisplay) {
                summaryAdvicePrintDisplay.textContent = (summaryAdviceTextarea ? summaryAdviceTextarea.value : '') || "アドバイスは未記入です。";
                summaryAdvicePrintDisplay.style.display = 'block';
            }
            if (summaryAdviceTextarea) summaryAdviceTextarea.style.display = 'none';

            html2pdf().set(opt).from(element).save().then(() => {
                if (summaryAdvicePrintDisplay) summaryAdvicePrintDisplay.style.display = 'none';
                if (summaryAdviceTextarea) summaryAdviceTextarea.style.display = 'block';
            }).catch(err => {
                console.error(err);
                if (summaryAdvicePrintDisplay) summaryAdvicePrintDisplay.style.display = 'none';
                if (summaryAdviceTextarea) summaryAdviceTextarea.style.display = 'block';
                alert('PDFの出力中にエラーが発生しました。');
            });
        });
    }

    // 4.1 Download active single day
    pdfSingleBtn.addEventListener('click', () => {
        // Save current advice first
        if (activeDateKey) {
            parsedData.days[activeDateKey].advice = adviceTextarea.value;
        }

        const element = document.getElementById('printable-report');
        const patientName = patientNameInput.value.trim();
        const d = new Date(activeDateKey);
        const dateStr = `${d.getMonth() + 1}月${d.getDate()}日`;
        
        const filename = `栄養食事指導参考資料_${patientName ? patientName + '様' : '食事記録'}_${dateStr}.pdf`;
        
        const opt = {
            margin: 10,
            filename: filename,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true, logging: false, letterRendering: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' },
            pagebreak: { mode: ['css', 'legacy'] }
        };

        // 一時的にアドバイス入力欄を隠し、印刷用プレビューDIVを表示（長いコメントの切れ防止）
        const advicePrintDisplay = document.getElementById('advice-print-display');
        if (advicePrintDisplay) {
            advicePrintDisplay.style.display = 'block';
        }
        adviceTextarea.style.display = 'none';

        html2pdf().set(opt).from(element).save().then(() => {
            // PDF出力完了後に元の入力画面表示に戻す
            if (advicePrintDisplay) {
                advicePrintDisplay.style.display = 'none';
            }
            adviceTextarea.style.display = 'block';
        });
    });

    // 4.2 Download all days combined
    pdfAllBtn.addEventListener('click', async () => {
        // Save current advice first
        if (activeDateKey) {
            parsedData.days[activeDateKey].advice = adviceTextarea.value;
        }

        const patientName = patientNameInput.value.trim();
        const filename = `栄養食事指導参考資料_${patientName ? patientName + '様' : '食事記録'}_全日程.pdf`;
        
        const printContainer = document.getElementById('all-days-print-container');
        printContainer.innerHTML = '';
        printContainer.style.display = 'block'; // Temporarily display for html2pdf to compile

        try {
            // Compile report HTML for each day sequentially
            for (let i = 0; i < parsedData.dates.length; i++) {
                const dateKey = parsedData.dates[i];
                const dayData = parsedData.days[dateKey];
                const dayIndex = i + 1;

                const dayLabel = dayData.dayLabel !== undefined ? dayData.dayLabel : String(dayIndex);

                // Create report wrapper
                const wrapper = document.createElement('div');
                wrapper.className = 'printable-report-wrapper';
                
                // Inject report structure template
                wrapper.innerHTML = createReportHtmlTemplate(dateKey, dayIndex, dayData.displayDate, patientName, dayData.advice, dayLabel);
                printContainer.appendChild(wrapper);

                // Populate dynamic tables for this date
                populateReportData(wrapper, dayData);

                // Render specific chart for this day page
                renderPfcChartForContainer(wrapper, dayData.totals, dayIndex);

                // 最後の日付でなければ、次の日付の前に強制改ページを挿入する
                if (i < parsedData.dates.length - 1) {
                    const pageBreak = document.createElement('div');
                    pageBreak.className = 'html2pdf__page-break';
                    printContainer.appendChild(pageBreak);
                }
            }

            const opt = {
                margin: 10,
                filename: filename,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true, logging: false, letterRendering: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' },
                pagebreak: { mode: ['css', 'legacy'] }
            };

            await html2pdf().set(opt).from(printContainer).save();

        } catch (err) {
            console.error(err);
            alert('一括PDFの出力中にエラーが発生しました。');
        } finally {
            printContainer.style.display = 'none';
            printContainer.innerHTML = ''; // Clean memory
        }
    });

    // Helper: Dynamic template creator for bulk print
    function createReportHtmlTemplate(dateKey, dayIndex, displayDate, patientName, advice, dayLabel) {
        return `
            <div class="printable-report" style="margin-bottom: 0px; border: none; box-shadow: none;">
                <!-- Header -->
                <div class="report-title-area">
                    <div class="title-main">
                        <h2>栄養食事指導参考資料</h2>
                        <p class="title-sub">食事記録レポート (${dayLabel}日目)</p>
                    </div>
                    <div class="meta-inputs">
                        <div class="meta-row">
                            <span class="meta-label">食事記録日:</span>
                            <span class="meta-value">${displayDate}</span>
                        </div>
                        <div class="meta-row">
                            <span class="meta-label">対象者様:</span>
                            <span class="meta-value" style="border-bottom: 1px solid #7f8c8d; padding: 2px 15px; font-weight:700; min-width:100px; text-align:center;">${patientName || '　　　　'}</span>
                            <span class="meta-suffix">様</span>
                        </div>
                    </div>
                </div>

                <!-- Meals Grid -->
                <div class="meals-grid">
                    ${['breakfast', 'lunch', 'dinner', 'other'].map(type => {
                        const jpName = type === 'breakfast' ? '朝食' : type === 'lunch' ? '昼食' : type === 'dinner' ? '夕食' : 'その他';
                        const emoji = type === 'breakfast' ? '☀️' : type === 'lunch' ? '🕛' : type === 'dinner' ? '🌙' : '☕';
                        return `
                            <div class="meal-card">
                                <div class="meal-card-left">
                                    <div class="meal-card-header bg-${type}">
                                        <span class="meal-icon">${emoji}</span>
                                        <h3>${jpName}</h3>
                                    </div>
                                    <div class="meal-images-container" id="images-${type}-${dayIndex}"></div>
                                </div>
                                <div class="meal-table-container">
                                    <table>
                                        <thead>
                                            <tr>
                                                <th>料理名</th>
                                                <th>メモ</th>
                                                <th style="width: 120px; text-align: right;">単位</th>
                                            </tr>
                                        </thead>
                                        <tbody id="list-${type}-${dayIndex}"></tbody>
                                    </table>
                                </div>
                            </div>
                        `;
                    }).join('')}
                </div>

                <!-- Lower Section Analysis -->
                <div class="analysis-section">
                    <div class="analysis-left">
                        <div class="section-subtitle">
                            <i data-lucide="bar-chart-3" class="subtitle-icon"></i>
                            <h3>栄養素摂取分析</h3>
                        </div>
                        <table class="analysis-table">
                            <thead>
                                <tr>
                                    <th>栄養区分</th>
                                    <th>エネルギー (kcal)</th>
                                    <th>たんぱく質 (g)</th>
                                    <th>脂質 (g)</th>
                                    <th>炭水化物 (g)</th>
                                    <th>食物繊維 (g)</th>
                                    <th>カルシウム (mg)</th>
                                    <th>鉄 (mg)</th>
                                    <th>塩分相当 (g)</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${['breakfast', 'lunch', 'dinner', 'other'].map((type, idx) => {
                                    const jpName = type === 'breakfast' ? '朝食' : type === 'lunch' ? '昼食' : type === 'dinner' ? '夕食' : 'その他';
                                    return `
                                        <tr>
                                            <td class="cell-label"><span class="dot bg-${type}"></span> ${jpName}</td>
                                            <td id="val-${type}-kcal-${dayIndex}" class="num-val">-</td>
                                            <td id="val-${type}-protein-${dayIndex}" class="num-val">-</td>
                                            <td id="val-${type}-fat-${dayIndex}" class="num-val">-</td>
                                            <td id="val-${type}-carb-${dayIndex}" class="num-val">-</td>
                                            ${idx === 0 ? `<td rowspan="4" class="disabled-cell text-muted">1日合計のみ<br>表示</td>
                                                           <td rowspan="4" class="disabled-cell text-muted">1日合計のみ<br>表示</td>
                                                           <td rowspan="4" class="disabled-cell text-muted">1日合計のみ<br>表示</td>` : ''}
                                            <td id="val-${type}-salt-${dayIndex}" class="num-val">-</td>
                                        </tr>
                                    `;
                                }).join('')}
                                <tr class="row-total">
                                    <td class="cell-label">1日合計</td>
                                    <td id="val-total-kcal-${dayIndex}" class="num-val">-</td>
                                    <td id="val-total-protein-${dayIndex}" class="num-val">-</td>
                                    <td id="val-total-fat-${dayIndex}" class="num-val">-</td>
                                    <td id="val-total-carb-${dayIndex}" class="num-val">-</td>
                                    <td id="val-total-fiber-${dayIndex}" class="num-val">-</td>
                                    <td id="val-total-calcium-${dayIndex}" class="num-val">-</td>
                                    <td id="val-total-iron-${dayIndex}" class="num-val">-</td>
                                    <td id="val-total-salt-${dayIndex}" class="num-val">-</td>
                                </tr>
                            </tbody>
                        </table>
                        
                        <div class="vitamins-grid">
                            <div class="vitamin-item"><span class="vit-label">ビタミンA:</span><span id="vit-a-${dayIndex}" class="vit-val">-</span></div>
                            <div class="vitamin-item"><span class="vit-label">ビタミンD:</span><span id="vit-d-${dayIndex}" class="vit-val">-</span></div>
                            <div class="vitamin-item"><span class="vit-label">ビタミンB1:</span><span id="vit-b1-${dayIndex}" class="vit-val">-</span></div>
                            <div class="vitamin-item"><span class="vit-label">ビタミンB2:</span><span id="vit-b2-${dayIndex}" class="vit-val">-</span></div>
                            <div class="vitamin-item"><span class="vit-label">ビタミンB6:</span><span id="vit-b6-${dayIndex}" class="vit-val">-</span></div>
                            <div class="vitamin-item"><span class="vit-label">ビタミンC:</span><span id="vit-c-${dayIndex}" class="vit-val">-</span></div>
                        </div>
                    </div>

                    <div class="analysis-right">
                        <div class="section-subtitle">
                            <i data-lucide="pie-chart" class="subtitle-icon"></i>
                            <h3>1日のPFCバランス</h3>
                        </div>
                        <div class="pfc-container">
                            <div class="chart-wrapper">
                                <canvas id="pfcChart-${dayIndex}"></canvas>
                            </div>
                            <div class="pfc-details">
                                <div class="pfc-legend-item protein"><span class="legend-dot bg-pfc-p"></span><span class="legend-name">P (たんぱく質)</span><span id="pfc-p-val-${dayIndex}" class="legend-val">-</span><span id="pfc-p-pct-${dayIndex}" class="legend-pct">-</span></div>
                                <div class="pfc-legend-item fat"><span class="legend-dot bg-pfc-f"></span><span class="legend-name">F (脂質)</span><span id="pfc-f-val-${dayIndex}" class="legend-val">-</span><span id="pfc-f-pct-${dayIndex}" class="legend-pct">-</span></div>
                                <div class="pfc-legend-item carb"><span class="legend-dot bg-pfc-c"></span><span class="legend-name">C (炭水化物)</span><span id="pfc-c-val-${dayIndex}" class="legend-val">-</span><span id="pfc-c-pct-${dayIndex}" class="legend-pct">-</span></div>
                                <div class="pfc-target-ref"><span class="ref-title">目標比率の目安（厚生労働省）:</span><span class="ref-body">P: 13〜20% | F: 20〜30% | C: 50〜65%</span></div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Advice Memo -->
                <div class="advice-section">
                    <div class="advice-header">
                        <i data-lucide="message-square" class="advice-icon"></i>
                        <span>栄養指導メモ・アドバイス欄</span>
                    </div>
                    <div class="advice-body">
                        <div class="advice-print-display" style="display: block;">${advice || 'アドバイスは未記入です。'}</div>
                    </div>
                </div>
            </div>
        `;
    }

    // Helper: Populate data values inside print wrapper
    function populateReportData(wrapper, dayData) {
        const dayIndex = wrapper.querySelector('.title-sub').textContent.match(/\((\d+)日目\)/)[1];
        const meals = dayData.meals;
        const totals = dayData.totals;

        const mealTypes = ['breakfast', 'lunch', 'dinner', 'other'];
        mealTypes.forEach(type => {
            const listEl = wrapper.querySelector(`#list-${type}-${dayIndex}`);
            const imgEl = wrapper.querySelector(`#images-${type}-${dayIndex}`);
            const data = meals[type];

            // Food List
            if (data.items.length === 0) {
                listEl.innerHTML = `<tr><td colspan="3" class="text-muted text-center" style="font-style:italic; padding:12px;">食事データなし</td></tr>`;
            } else {
                data.items.forEach(item => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `<td style="font-weight: 500;">${item.name}</td><td class="text-muted">${item.memo}</td><td style="text-align: right; font-weight: 500;">${item.unit}</td>`;
                    listEl.appendChild(tr);
                });
            }

            // Food Images
            if (data.images.length === 0) {
                imgEl.innerHTML = `<span class="no-image-placeholder">写真なし</span>`;
            } else {
                data.images.forEach(imgSrc => {
                    const wrapperDiv = document.createElement('div');
                    wrapperDiv.className = 'meal-img-wrapper';
                    const img = document.createElement('img');
                    img.src = imgSrc;
                    img.className = 'meal-img';
                    img.alt = `${type} image`;
                    wrapperDiv.appendChild(img);
                    imgEl.appendChild(wrapperDiv);
                });
            }

            // Sub-nutrition values
            wrapper.querySelector(`#val-${type}-kcal-${dayIndex}`).textContent = `${Math.round(data.kcal)} kcal`;
            wrapper.querySelector(`#val-${type}-protein-${dayIndex}`).textContent = `${data.protein.toFixed(1)} g`;
            wrapper.querySelector(`#val-${type}-fat-${dayIndex}`).textContent = `${data.fat.toFixed(1)} g`;
            wrapper.querySelector(`#val-${type}-carb-${dayIndex}`).textContent = `${data.carb.toFixed(1)} g`;
            wrapper.querySelector(`#val-${type}-salt-${dayIndex}`).textContent = `${data.salt.toFixed(1)} g`;
        });

        // Totals
        wrapper.querySelector(`#val-total-kcal-${dayIndex}`).textContent = `${Math.round(totals.kcal)} kcal`;
        wrapper.querySelector(`#val-total-protein-${dayIndex}`).textContent = `${totals.protein.toFixed(1)} g`;
        wrapper.querySelector(`#val-total-fat-${dayIndex}`).textContent = `${totals.fat.toFixed(1)} g`;
        wrapper.querySelector(`#val-total-carb-${dayIndex}`).textContent = `${totals.carb.toFixed(1)} g`;
        wrapper.querySelector(`#val-total-fiber-${dayIndex}`).textContent = `${totals.fiber.toFixed(1)} g`;
        wrapper.querySelector(`#val-total-calcium-${dayIndex}`).textContent = `${Math.round(totals.calcium)} mg`;
        wrapper.querySelector(`#val-total-iron-${dayIndex}`).textContent = `${totals.iron.toFixed(1)} mg`;
        wrapper.querySelector(`#val-total-salt-${dayIndex}`).textContent = `${totals.salt.toFixed(1)} g`;

        // Extra vitamins
        wrapper.querySelector(`#vit-a-${dayIndex}`).textContent = `${Math.round(totals.vitA)}`;
        wrapper.querySelector(`#vit-d-${dayIndex}`).textContent = `${totals.vitD.toFixed(1)}`;
        wrapper.querySelector(`#vit-b1-${dayIndex}`).textContent = `${totals.vitB1.toFixed(2)}`;
        wrapper.querySelector(`#vit-b2-${dayIndex}`).textContent = `${totals.vitB2.toFixed(2)}`;
        wrapper.querySelector(`#vit-b6-${dayIndex}`).textContent = `${totals.vitB6.toFixed(2)}`;
        wrapper.querySelector(`#vit-c-${dayIndex}`).textContent = `${Math.round(totals.vitC)}`;
    }

    // Helper: Render PFC chart on bulk page
    function renderPfcChartForContainer(wrapper, totals, dayIndex) {
        const pKcal = totals.protein * 4;
        const fKcal = totals.fat * 9;
        const cKcal = totals.carb * 4;
        const totalPfcKcal = pKcal + fKcal + cKcal;

        let pPct = 0, fPct = 0, cPct = 0;
        if (totalPfcKcal > 0) {
            pPct = (pKcal / totalPfcKcal) * 100;
            fPct = (fKcal / totalPfcKcal) * 100;
            cPct = (cKcal / totalPfcKcal) * 100;
        }

        // Legends
        wrapper.querySelector(`#pfc-p-val-${dayIndex}`).textContent = `${totals.protein.toFixed(1)}g (${Math.round(pKcal)} kcal)`;
        wrapper.querySelector(`#pfc-p-pct-${dayIndex}`).textContent = `${pPct.toFixed(1)}%`;
        
        wrapper.querySelector(`#pfc-f-val-${dayIndex}`).textContent = `${totals.fat.toFixed(1)}g (${Math.round(fKcal)} kcal)`;
        wrapper.querySelector(`#pfc-f-pct-${dayIndex}`).textContent = `${fPct.toFixed(1)}%`;
        
        wrapper.querySelector(`#pfc-c-val-${dayIndex}`).textContent = `${totals.carb.toFixed(1)}g (${Math.round(cKcal)} kcal)`;
        wrapper.querySelector(`#pfc-c-pct-${dayIndex}`).textContent = `${cPct.toFixed(1)}%`;

        const ctx = wrapper.querySelector(`#pfcChart-${dayIndex}`).getContext('2d');
        new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['P (たんぱく質)', 'F (脂質)', 'C (炭水化物)'],
                datasets: [{
                    data: [pKcal, fKcal, cKcal],
                    backgroundColor: ['#ff6b6b', '#feca57', '#48dbfb'],
                    borderWidth: 1.5,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                plugins: {
                    legend: { display: false },
                    tooltip: { enabled: false } // No tooltips needed on static PDF output
                },
                cutout: '65%',
                responsive: true,
                maintainAspectRatio: false
            }
        });
    }

    backBtn.addEventListener('click', () => {
        showScreen('upload-section');
    });

    // ==========================================
    // 5. VIEW TRANSITIONS HELPERS
    // ==========================================
    function showScreen(screenId) {
        document.querySelectorAll('.screen-section').forEach(section => {
            section.classList.remove('active');
        });
        
        const targetSection = document.getElementById(screenId);
        targetSection.classList.add('active');
        lucide.createIcons();
    }

    function showLoading(isLoading) {
        if (isLoading) {
            dropZone.style.display = 'none';
            parseLoading.style.display = 'flex';
        } else {
            dropZone.style.display = 'flex';
            parseLoading.style.display = 'none';
        }
    }

    function showUploadError(message) {
        if (message) {
            uploadError.textContent = message;
            uploadError.style.display = 'flex';
        } else {
            uploadError.textContent = '';
            uploadError.style.display = 'none';
        }
    }
});
