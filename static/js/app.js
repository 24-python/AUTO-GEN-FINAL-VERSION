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
            error: null
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
        async processFiles() {
            if (this.selectedFiles.length === 0) return;

            this.isProcessing = true;
            this.progress = 0;
            this.progressMessage = `Обработка ${this.selectedFiles.length} файл(ов)...`;
            this.result = null;
            this.error = null;

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
                const response = await fetch('/api/generate', {
                    method: 'POST',
                    body: formData
                });
                const data = await response.json();

                if (data.success) {
                    this.progress = 100;
                    this.progressMessage = 'Готово!';
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
                this.error = 'Ошибка соединения с сервером';
                Toastify({ text: '❌ ' + this.error, duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            } finally {
                this.isProcessing = false;
            }
        }
    }
}).mount('#app');