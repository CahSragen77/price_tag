 let masterData = [];
        let printQueue = [];
        let isOnline = true;
        let html5QrcodeScanner = null;

        // Load Konfigurasi URL Google Sheets dari Browser Storage jika ada
        window.onload = function() {
            if (localStorage.getItem('sheetsUrl')) {
                document.getElementById('sheetsUrlInput').value = localStorage.getItem('sheetsUrl');
            }
            checkNetworkStatus();
            fetchOnlineData();
        };

        function saveConfig() {
            localStorage.setItem('sheetsUrl', document.getElementById('sheetsUrlInput').value.trim());
        }

        // --- Cek Status Jaringan ---
        function checkNetworkStatus() {
            const statusBadge = document.getElementById('connectionStatus');
            const btnUpload = document.getElementById('btnUploadManual');
            
            if (navigator.onLine) {
                isOnline = true;
                statusBadge.innerText = "Status: Online (Terhubung Internet)";
                statusBadge.className = "status-badge online";
                if(masterData.length > 0) {
                    btnUpload.style.display = 'inline-block';
                }
            } else {
                isOnline = false;
                statusBadge.innerText = "Status: Offline (Koneksi Terputus)";
                statusBadge.className = "status-badge offline";
                btnUpload.style.display = 'none';
                console.log("Koneksi internet terputus. Sistem otomatis menggunakan data lokal offline.");
            }
        }

        window.addEventListener('online', checkNetworkStatus);
        window.addEventListener('offline', checkNetworkStatus);

        // --- Ambil Data dari Google Sheet secara Online ---
        function fetchOnlineData() {
            checkNetworkStatus();
            const url = document.getElementById('sheetsUrlInput').value.trim();
            if (!url) return;

            if (!isOnline) {
                alert("Anda sedang offline. Tidak dapat melakukan sinkronisasi dengan Google Sheet.");
                return;
            }

            fetch(url + "?action=read")
            .then(response => response.json())
            .then(data => {
                if(data && data.length > 0) {
                    masterData = data.map(item => ({
                        barcode: String(item.barcode).trim(),
                        desc: String(item.desc).trim(),
                        price: parseFloat(item.price) || 0,
                        unit: String(item.unit || 'PCS').trim()
                    }));
                    alert(`Berhasil sinkronisasi! ${masterData.length} data dimuat dari Google Sheet secara Online.`);
                }
            })
            .catch(err => {
                console.error("Gagal mengambil data dari Google Sheet: ", err);
                alert("Gagal sinkronisasi dengan Google Sheet. Pastikan URL Web App benar dan sudah di-Deploy sebagai 'Anyone'.");
            });
        }

        // --- Upload Manual Data Lokal ke Google Sheet ---
        function uploadLocalToOnline() {
            checkNetworkStatus();
            const url = document.getElementById('sheetsUrlInput').value.trim();
            if (!url) return alert("Masukkan URL Web App Google Sheet terlebih dahulu!");
            if (!isOnline) return alert("Tidak ada koneksi internet untuk mengunggah!");
            if (masterData.length === 0) return alert("Tidak ada data lokal yang bisa diupload.");

            if(confirm(`Apakah Anda ingin mengunggah ${masterData.length} data lokal saat ini ke Google Sheet?`)) {
                fetch(url, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ action: 'write', data: masterData })
                })
                .then(() => {
                    alert("Proses kirim data lokal selesai! Silakan periksa Google Sheet Anda.");
                })
                .catch(err => {
                    alert("Gagal mengunggah data lokal.");
                    console.error(err);
                });
            }
        }

        // --- FITUR BARU: CAMERA SCANNER HP (KAMERA BELAKANG) ---
        function toggleScanner() {
            const container = document.getElementById('scanner-container');
            if (container.style.display === 'none' || container.style.display === '') {
                container.style.display = 'block';
                
                html5QrcodeScanner = new Html5Qrcode("interactive-reader");
                html5QrcodeScanner.start(
                    { facingMode: "environment" }, // Paksa Kamera Belakang HP
                    {
                        fps: 10,
                        qrbox: { width: 250, height: 150 }
                    },
                    (decodedText) => {
                        // Jika barcode terdeteksi
                        document.getElementById('manualBarcode').value = decodedText;
                        searchBarcode();
                        
                        // Efek Audio Bip Ringan saat berhasil scan
                        try {
                            let context = new (window.AudioContext || window.webkitAudioContext)();
                            let osc = context.createOscillator();
                            osc.type = "sine";
                            osc.frequency.value = 1200;
                            osc.connect(context.destination);
                            osc.start();
                            setTimeout(() => osc.stop(), 100);
                        } catch(e){}
                    },
                    (errorMessage) => { /* Silently ignore error scanning */ }
                ).catch(err => {
                    alert("Gagal mengakses kamera belakang.");
                    console.error(err);
                    container.style.display = 'none';
                });
            } else {
                container.style.display = 'none';
                if (html5QrcodeScanner) {
                    html5QrcodeScanner.stop().then(() => {
                        html5QrcodeScanner = null;
                    });
                }
            }
        }

        // --- 1. Baca Data Master (Excel Harga - OFFLINE BACKUP) ---
        document.getElementById('fileInput').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(sheet);
                
                masterData = [];
                rows.forEach(row => {
                    let barcode = row['Barcode'] || row['BARCODE'];
                    let desc = row['Desc Name'] || row['NAMA BARANG'];
                    let unit = row['Unit'] || row['Satuan'] || 'PCS';
                    let price = 0;
                    
                    if (row['HJ-MM'] && row['HJ-MM'] !== 0) {
                        price = row['HJ-MM'];
                    } else if (row['NEW RETAIL PRICE HJ MM']) {
                        price = row['NEW RETAIL PRICE HJ MM'];
                    }

                    if (barcode && desc) {
                        masterData.push({
                            barcode: String(barcode).trim(),
                            desc: String(desc).trim(),
                            price: parseFloat(price) || 0,
                            unit: String(unit).trim()
                        });
                    }
                });
                alert(`Berhasil memuat ${masterData.length} data dari Excel Master Lokal!`);
                checkNetworkStatus();
            };
            reader.readAsArrayBuffer(file);
        });

        // --- 2. Upload File LoMag ---
        document.getElementById('listBarcodeInput').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    const firstSheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[firstSheetName];
                    const rows = XLSX.utils.sheet_to_json(sheet);
                    
                    let addedCount = 0;
                    rows.forEach(row => {
                        let bc = row['Barcode'] || row['BARCODE'] || row['Item code'] || row['Item Code'];
                        
                        if (!bc) {
                            const keys = Object.keys(row);
                            if (keys.length > 0) {
                                bc = row[keys[0]];
                            }
                        }

                        if (bc) {
                            const item = findItem(String(bc).trim());
                            if (item) {
                                printQueue.push(item);
                                addedCount++;
                            }
                        }
                    });
                    
                    updateQueueDisplay();
                    alert(`Berhasil menambahkan ${addedCount} item dari file Excel LoMag.`);
                } catch (err) {
                    alert("Gagal membaca file Excel. Pastikan formatnya .xls/.xlsx.");
                    console.error(err);
                }
            };
            reader.readAsArrayBuffer(file);
        });

        // --- 3. Input Manual ---
        function searchBarcode() {
            const input = document.getElementById('manualBarcode');
            const bc = input.value.trim();
            if (!bc) return alert('Masukkan Barcode!');
            const item = findItem(bc);
            if (item) {
                printQueue.push(item);
                updateQueueDisplay();
                input.value = '';
            } else {
                alert(`Barcode ${bc} tidak ditemukan di Data Master!`);
            }
        }

        function findItem(barcode) { return masterData.find(item => item.barcode === barcode); }

        function updateQueueDisplay() {
            const queue = document.getElementById('queueList');
            queue.value = printQueue.map(item => `${item.barcode} - ${item.desc}`).join('\n');
        }

        function clearQueue() {
            printQueue = [];
            updateQueueDisplay();
            document.getElementById('printArea').innerHTML = '';
        }

        // --- 4. Generate Print ---
        function preparePrint() {
            if (printQueue.length === 0) return alert('Antrian kosong!');
            if (masterData.length === 0) return alert('Data Excel Master belum diload!');

            const size = document.getElementById('sizeSelect').value;
            const expiryText = document.getElementById('expiryInput').value.trim() || 'R_H-3 MONTH';
            const today = new Date();
            const dateStr = today.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-');

            let htmlContent = `<div class="pricetag-wrapper">`;

            printQueue.forEach(item => {
                const className = size === 'lg' ? 'pricetag-lg' : 'pricetag-sm';
                const priceStr = `Rp. ${item.price.toLocaleString('id-ID')}`;

                htmlContent += `
                    <div class="${className}">
                        <div class="p-top ${size === 'sm' ? 'p-top-sm' : ''}">
                            ${item.desc}
                        </div>
                        <div class="p-harga ${size === 'sm' ? 'p-harga-sm' : ''}">
                            ${priceStr} <span class="p-satuan ${size === 'sm' ? 'p-satuan-sm' : ''}">/ ${item.unit}</span>
                        </div>
                        
                        <div class="p-barcode-area">
                            <div class="p-barcode-wrap">
                                <svg class="barcode-js" data-barcode="${item.barcode}"></svg>
                                <div class="p-barcode-text">${item.barcode}</div>
                            </div>
                            ${size === 'lg' ? `
                            <div class="p-logo-area">
                                <svg class="p-logo-svg" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
                                    <circle cx="50" cy="50" r="45" stroke="#00a2e8" stroke-width="5" fill="white"/>
                                    <path d="M30 40 H70 V45 H55 V75 H45 V45 H30 Z" fill="#00a2e8"/> 
                                    <rect x="30" y="55" width="40" height="10" fill="#00a2e8"/>
                                </svg>
                                <div class="p-logo-text">amanda<br>mart</div>
                            </div>
                            ` : ''}
                        </div>

                        <div class="p-strip-container">
                            <div class="p-bottom-text">
                                <div class="p-left">${expiryText}</div>
                                <div class="p-date">${dateStr}</div>
                            </div>
                            <div class="p-line-blue"></div>
                            <div class="p-gap"></div>
                            <div class="p-line-red"></div>
                        </div>
                    </div>
                `;
            });

            htmlContent += `</div>`;
            
            const printArea = document.getElementById('printArea');
            printArea.innerHTML = htmlContent;

            setTimeout(() => {
                const svgs = printArea.querySelectorAll('.barcode-js');
                svgs.forEach(svg => {
                    const barcodeValue = svg.getAttribute('data-barcode');
                    try {
                        JsBarcode(svg, barcodeValue, {
                            format: "CODE128",
                            width: 2,
                            height: 28,
                            displayValue: false,
                            margin: 0
                        });
                    } catch (err) {
                        console.warn("Gagal barcode:", barcodeValue);
                    }
                });
                
                window.print();
            }, 500);
        }
