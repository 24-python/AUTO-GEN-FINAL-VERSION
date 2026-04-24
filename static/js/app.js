const { createApp } = Vue;

createApp({
    data() {
        return {
            selectedFiles: [],
            mode: 1,
            isProcessing: false,
            progress: 0,
            progressMessage: '',
            result: null,
            error: null,
            progressInterval: null
        }
    },
    methods: {
        triggerUpload() {
            document.getElementById('fileInput').click();
        },
        addMoreFiles() {
            document.getElementById('fileInput').click();
        },
        handleFileSelect(e) {
            const files = Array.from(e.target.files);
            const validFiles = files.filter(f => {
                const ext = f.name.split('.').pop().toLowerCase();
                return ext === 'docx' || ext === 'zip';
            });
            if (validFiles.length < files.length) {
                Toastify({ text: '⚠️ Пропущены файлы неверного формата', duration: 3000, gravity: 'bottom', position: 'right' }).showToast();
            }
            this.selectedFiles = [...this.selectedFiles, ...validFiles];
            this.result = null;
            this.error = null;
            e.target.value = '';
        },
        handleDrop(e) {
            const files = Array.from(e.dataTransfer.files);
            const validFiles = files.filter(f => {
                const ext = f.name.split('.').pop().toLowerCase();
                return ext === 'docx' || ext === 'zip';
            });
            if (validFiles.length < files.length) {
                Toastify({ text: '⚠️ Пропущены файлы неверного формата', duration: 3000, gravity: 'bottom', position: 'right' }).showToast();
            }
            this.selectedFiles = [...this.selectedFiles, ...validFiles];
            this.result = null;
            this.error = null;
        },
        removeFile(idx) {
            this.selectedFiles.splice(idx, 1);
            this.result = null;
            this.error = null;
        },
        resetAll() {
            this.selectedFiles = [];
            this.result = null;
            this.error = null;
            this.progress = 0;
            this.progressMessage = '';
            document.getElementById('fileInput').value = '';
            if (this.progressInterval) {
                clearInterval(this.progressInterval);
                this.progressInterval = null;
            }
        },
        downloadZip() {
            if (this.result && this.result.zip_url) {
                window.location.href = this.result.zip_url;
            }
        },
        startProgressSimulation() {
            this.progress = 0;
            const steps = [
                { pct: 15, msg: 'Парсинг чек-листа...' },
                { pct: 35, msg: 'Загрузка инструкций из БД...' },
                { pct: 55, msg: 'Группировка объектов...' },
                { pct: 75, msg: 'Заполнение таблицы Word...' },
                { pct: 90, msg: 'Сохранение документа...' },
            ];

            let stepIndex = 0;
            this.progressMessage = steps[0].msg;

            this.progressInterval = setInterval(() => {
                if (stepIndex < steps.length) {
                    const targetPct = steps[stepIndex].pct;
                    const targetMsg = steps[stepIndex].msg;

                    // Плавно увеличиваем прогресс
                    const inc = (targetPct - this.progress) / 10;
                    const smoothInterval = setInterval(() => {
                        if (this.progress < targetPct) {
                            this.progress = Math.min(this.progress + inc, targetPct);
                        } else {
                            clearInterval(smoothInterval);
                        }
                    }, 100);

                    this.progressMessage = targetMsg;
                    stepIndex++;
                }
            }, 800);
        },
        stopProgressSimulation(success = true) {
            if (this.progressInterval) {
                clearInterval(this.progressInterval);
                this.progressInterval = null;
            }

            if (success) {
                this.progress = 100;
                this.progressMessage = '✅ Генерация завершена!';
            } else {
                this.progressMessage = '❌ Ошибка генерации';
            }
        },
        async processFiles() {
            if (this.selectedFiles.length === 0) return;

            this.isProcessing = true;
            this.result = null;
            this.error = null;

            // Запускаем симуляцию прогресса
            this.startProgressSimulation();

            const formData = new FormData();
            if (this.selectedFiles.length === 1) {
                formData.append('file', this.selectedFiles[0]);
            } else {
                formData.append('file', this.selectedFiles[0]);
                for (let i = 1; i < this.selectedFiles.length; i++) {
                    formData.append('file_' + i, this.selectedFiles[i]);
                }
                formData.append('multi', 'true');
            }
            formData.append('mode', this.mode);

            try {
                const response = await fetch('/api/generate', { method: 'POST', body: formData });
                const data = await response.json();

                // Останавливаем симуляцию прогресса
                this.stopProgressSimulation(data.success);

                if (data.success) {
                    this.result = data;
                    const msg = data.total_files > 1
                        ? `✅ Обработано файлов: ${data.total_files}, объектов: ${data.total_objects}`
                        : `✅ Техкарта сгенерирована (${data.total_objects} объектов)`;
                    Toastify({ text: msg, duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                } else {
                    this.error = data.error || 'Ошибка при обработке файла';
                    Toastify({ text: '❌ ' + this.error, duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                this.stopProgressSimulation(false);
                this.error = 'Ошибка соединения с сервером';
                Toastify({ text: '❌ ' + this.error, duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            } finally {
                this.isProcessing = false;
            }
        }
    },
    beforeUnmount() {
        if (this.progressInterval) {
            clearInterval(this.progressInterval);
        }
    }
}).mount('#app');