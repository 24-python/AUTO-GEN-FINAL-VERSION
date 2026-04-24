const { createApp } = Vue;

createApp({
    data() {
        return {
            objects: [],
            showImportModal: false,
            searchQuery: ''
        }
    },
    mounted() {
        this.loadObjects();
    },
    methods: {
        async loadObjects() {
            try {
                const res = await fetch('/api/objects');
                const data = await res.json();
                this.objects = data.objects || [];
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки данных', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async loadCSV() {
            window.location.href = '/api/export_csv';
        },
        editObject(obj) {
            Toastify({ text: 'Редактирование: ' + obj.display_name, duration: 2000, gravity: 'bottom', position: 'right' }).showToast();
        },
        async deleteObject(obj) {
            if (!confirm('Удалить объект «' + obj.display_name + '»?')) return;
            try {
                const res = await fetch('/api/objects/' + obj.id, { method: 'DELETE' });
                if (res.ok) {
                    await this.loadObjects();
                    Toastify({ text: '✅ Объект удалён', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        }
    }
}).mount('#adminApp');