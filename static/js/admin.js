const { createApp } = Vue;

createApp({
    data() {
        return {
            objects: [],
            categories: [],
            searchQuery: '',
            sortField: 'sort_priority',
            sortDir: 'asc',

            showEditModal: false,
            showAddModal: false,
            showImportModal: false,
            showInstructionsModal: false,

            editObject: null,
            formData: {
                display_name: '',
                normalized_name: '',
                sort_priority: 0,
                category_id: 1,
                base_name: '',
                modifier: ''
            },

            currentInstructions: [],

            importMode: 1,
            importFile: null
        }
    },
    computed: {
        filteredObjects() {
            let result = [...this.objects];
            if (this.searchQuery) {
                const q = this.searchQuery.toLowerCase();
                result = result.filter(o =>
                    (o.display_name || '').toLowerCase().includes(q) ||
                    (o.category_name || '').toLowerCase().includes(q) ||
                    (o.normalized_name || '').toLowerCase().includes(q)
                );
            }
            result.sort((a, b) => {
                let valA = a[this.sortField] || '';
                let valB = b[this.sortField] || '';
                if (typeof valA === 'string') valA = valA.toLowerCase();
                if (typeof valB === 'string') valB = valB.toLowerCase();
                if (valA < valB) return this.sortDir === 'asc' ? -1 : 1;
                if (valA > valB) return this.sortDir === 'asc' ? 1 : -1;
                return 0;
            });
            return result;
        }
    },
    mounted() {
        this.loadObjects();
        this.loadCategories();
    },
    methods: {
        async loadObjects() {
            try {
                const res = await fetch('/api/objects');
                const data = await res.json();
                this.objects = data.objects || [];
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async loadCategories() {
            try {
                const res = await fetch('/api/categories');
                const data = await res.json();
                this.categories = data.categories || [];
            } catch (e) {}
        },

        toggleSort(field) {
            if (this.sortField === field) {
                this.sortDir = this.sortDir === 'asc' ? 'desc' : 'asc';
            } else {
                this.sortField = field;
                this.sortDir = 'asc';
            }
        },

        openEditModal(obj) {
            this.editObject = obj;
            this.formData = {
                display_name: obj.display_name,
                normalized_name: obj.normalized_name,
                sort_priority: obj.sort_priority,
                category_id: obj.category_id || 1,
                base_name: obj.base_name || obj.display_name,
                modifier: obj.modifier || ''
            };
            this.showEditModal = true;
        },
        async saveEdit() {
            if (!this.editObject) return;
            try {
                const res = await fetch(`/api/objects/${this.editObject.id}`, {
                    method: 'PUT',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(this.formData)
                });
                if (res.ok) {
                    Toastify({ text: '✅ Объект обновлён', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showEditModal = false;
                    await this.loadObjects();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка сохранения', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        openAddModal() {
            this.formData = { display_name: '', normalized_name: '', sort_priority: 0, category_id: 1, base_name: '', modifier: '' };
            this.showAddModal = true;
        },
        async saveAdd() {
            try {
                const res = await fetch('/api/objects', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(this.formData)
                });
                if (res.ok) {
                    Toastify({ text: '✅ Объект создан', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showAddModal = false;
                    await this.loadObjects();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка создания', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        async deleteObject(obj) {
            if (!confirm(`Удалить «${obj.display_name}»?`)) return;
            try {
                await fetch(`/api/objects/${obj.id}`, { method: 'DELETE' });
                Toastify({ text: '✅ Объект удалён', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                await this.loadObjects();
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        async openInstructions(obj) {
            try {
                const res = await fetch(`/api/objects/${obj.id}/instructions`);
                const data = await res.json();
                this.currentInstructions = (data.instructions || []).map(i => ({...i, isNew: false}));
                this.editObject = obj;
                this.showInstructionsModal = true;
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки инструкций', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        addInstruction() {
            this.currentInstructions.push({
                id: null, isNew: true,
                cleaning_method: '', product_name: '', cleaning_technique: '',
                concentration: '', temperature: '', exposure_time: '',
                inventory: '', frequency: '', executor: '', control_method: '',
                instruction_number: ''
            });
        },
        async saveInstruction(instr) {
            try {
                if (instr.isNew) {
                    const res = await fetch(`/api/objects/${this.editObject.id}/instructions`, {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(instr)
                    });
                    const data = await res.json();
                    instr.id = data.id;
                    instr.isNew = false;
                    Toastify({ text: '✅ Инструкция добавлена', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                } else {
                    await fetch(`/api/instructions/${instr.id}`, {
                        method: 'PUT',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(instr)
                    });
                    Toastify({ text: '✅ Инструкция обновлена', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка сохранения', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async deleteInstruction(instr, idx) {
            if (!confirm('Удалить эту инструкцию?')) return;
            if (instr.isNew) {
                this.currentInstructions.splice(idx, 1);
                return;
            }
            try {
                await fetch(`/api/instructions/${instr.id}`, { method: 'DELETE' });
                this.currentInstructions.splice(idx, 1);
                Toastify({ text: '✅ Инструкция удалена', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        handleImportFile(e) { this.importFile = e.target.files[0]; },
        async doImport() {
            if (!this.importFile) return;
            const formData = new FormData();
            formData.append('file', this.importFile);
            formData.append('mode', this.importMode);
            try {
                const res = await fetch('/api/import_csv', { method: 'POST', body: formData });
                const data = await res.json();
                if (data.success) {
                    Toastify({ text: `✅ Импорт: создано ${data.created}, обновлено ${data.updated}`, duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showImportModal = false;
                    await this.loadObjects();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка импорта', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        }
    }
}).mount('#adminApp');