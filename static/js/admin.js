const { createApp } = Vue;

createApp({
    data() {
        return {
            objects: [],
            categories: [],
            roomCategories: [],
            products: [],                     // средства
            searchQuery: '',
            sortField: 'sort_priority',
            sortDir: 'asc',

            showEditModal: false,
            showAddModal: false,
            showImportModal: false,
            showInstructionsModal: false,
            showProductForm: false,           // форма средства
            showProductImportModal: false,    // импорт средств CSV
            editingProduct: null,            // редактируемое средство
            productForm: { name: '', product_type: '', color: '' },  // поля средства
            productImportFile: null,          // файл для импорта средств

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
            importFile: null,

            showRoomCatForm: false,
            editingRoomCat: null,
            roomCatForm: { name: '' },
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
        this.loadRoomCategories();
        this.loadProducts();                // загрузка средств
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
        async loadRoomCategories() {
            try {
                const res = await fetch('/api/room-categories');
                const data = await res.json();
                this.roomCategories = data.room_categories || [];
            } catch (e) {}
        },
        async loadProducts() {
            try {
                const res = await fetch('/api/products');
                const data = await res.json();
                this.products = data.products || [];
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

        // ========== ОЧИСТКА БАЗЫ ДАННЫХ ==========
        async clearDatabase() {
            if (!confirm('⚠️ ВНИМАНИЕ! Все объекты и инструкции будут удалены без возможности восстановления. Продолжить?')) return;
            if (!confirm('Точно удалить ВСЕ данные?')) return;

            try {
                const res = await fetch('/api/objects');
                const data = await res.json();
                const allObjects = data.objects || [];

                for (const obj of allObjects) {
                    await fetch(`/api/objects/${obj.id}`, { method: 'DELETE' });
                }

                Toastify({ text: `✅ База данных очищена (${allObjects.length} объектов)`, duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                await this.loadObjects();
                await this.loadRoomCategories();
            } catch (e) {
                Toastify({ text: '❌ Ошибка очистки БД', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        // ========== СРЕДСТВА ==========
        openProductForm() {
            this.editingProduct = null;
            this.productForm = { name: '', product_type: '', color: '' };
            this.showProductForm = true;
        },
        editProduct(p) {
            this.editingProduct = p;
            this.productForm = { name: p.name, product_type: p.product_type || '', color: p.color || '' };
            this.showProductForm = true;
        },
        async saveProduct() {
            const payload = {
                name: this.productForm.name.trim(),
                product_type: this.productForm.product_type,
                color: this.productForm.color.trim()
            };
            if (!payload.name) {
                Toastify({ text: '⚠️ Введите название', duration: 3000, gravity: 'bottom', position: 'right' }).showToast();
                return;
            }
            try {
                const url = this.editingProduct ? `/api/products/${this.editingProduct.id}` : '/api/products';
                const method = this.editingProduct ? 'PUT' : 'POST';
                const res = await fetch(url, {
                    method,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const data = await res.json();
                if (res.ok) {
                    Toastify({ text: '✅ Средство сохранено', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showProductForm = false;
                    this.editingProduct = null;
                    await this.loadProducts();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async deleteProduct(p) {
            if (!confirm(`Удалить средство «${p.name}»?`)) return;
            try {
                await fetch(`/api/products/${p.id}`, { method: 'DELETE' });
                Toastify({ text: '✅ Средство удалено', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                await this.loadProducts();
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        // Импорт/экспорт средств
        openProductImportModal() {
            this.productImportFile = null;
            this.showProductImportModal = true;
        },
        handleProductImportFile(e) {
            this.productImportFile = e.target.files[0];
        },
        async doProductImport() {
            if (!this.productImportFile) return;
            const formData = new FormData();
            formData.append('file', this.productImportFile);
            try {
                const res = await fetch('/api/products/import_csv', { method: 'POST', body: formData });
                const data = await res.json();
                if (data.success) {
                    Toastify({ text: `✅ Импорт: создано ${data.created || 0}, обновлено ${data.updated || 0}`, duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showProductImportModal = false;
                    await this.loadProducts();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        openRoomCatForm() {
            this.editingRoomCat = null;
            this.roomCatForm.name = '';
            this.showRoomCatForm = true;
        },
        editRoomCategory(rc) {
            this.editingRoomCat = rc;
            this.roomCatForm.name = rc.name;
            this.showRoomCatForm = true;
        },
        async saveRoomCategory() {
            const name = this.roomCatForm.name.trim();
            if (!name) {
                Toastify({ text: '⚠️ Введите название', duration: 3000, gravity: 'bottom', position: 'right' }).showToast();
                return;
            }
            try {
                const url = this.editingRoomCat
                    ? `/api/room-categories/${this.editingRoomCat.id}`
                    : '/api/room-categories';
                const method = this.editingRoomCat ? 'PUT' : 'POST';
                const res = await fetch(url, {
                    method,
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ name })
                });
                const data = await res.json();
                if (res.ok) {
                    Toastify({ text: '✅ Категория сохранена', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showRoomCatForm = false;
                    this.editingRoomCat = null;
                    this.roomCatForm.name = '';
                    await this.loadRoomCategories();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async deleteRoomCategory(rc) {
            if (!confirm(`Удалить категорию «${rc.name}»?`)) return;
            try {
                await fetch(`/api/room-categories/${rc.id}`, { method: 'DELETE' });
                Toastify({ text: '✅ Категория удалена', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                await this.loadRoomCategories();
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
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
                this.currentInstructions = (data.instructions || []).map(i => ({
                    ...i,
                    isNew: false,
                    room_category_id: i.room_category_id || null,
                    maintenance_type: i.maintenance_type || 'основная',
                    surface_type: i.surface_type || null
                }));
                this.editObject = obj;
                this.showInstructionsModal = true;
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки инструкций', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        aaddInstruction() {
    this.currentInstructions.push({
        id: null, isNew: true,
        room_category_id: null,
        maintenance_type: 'основная',
        cleaning_method: '', product_name: '', cleaning_technique: '',
        concentration: '', application_method: '', temperature: '', exposure_time: '',
        inventory: '', frequency: '', executor: '', control_method: '',
        instruction_number: '',
        surface_type: null
    });
},
        duplicateInstruction(instr, idx) {
            const copy = JSON.parse(JSON.stringify(instr));
            copy.id = null;
            copy.isNew = true;
            this.currentInstructions.splice(idx + 1, 0, copy);
            Toastify({ text: '✅ Инструкция дублирована', duration: 2000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
        },
        async saveInstruction(instr) {
            try {
                const payload = {
    room_category_id: instr.room_category_id || null,
    maintenance_type: instr.maintenance_type,
    cleaning_method: instr.cleaning_method,
    product_name: instr.product_name,
    cleaning_technique: instr.cleaning_technique,
    concentration: instr.concentration,
    application_method: instr.application_method || '',   // <--
    temperature: instr.temperature,
    exposure_time: instr.exposure_time,
    inventory: instr.inventory,
    frequency: instr.frequency,
    executor: instr.executor,
    control_method: instr.control_method,
    instruction_number: instr.instruction_number,
    surface_type: instr.surface_type || null
};

                if (instr.isNew) {
                    const res = await fetch(`/api/objects/${this.editObject.id}/instructions`, {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(payload)
                    });
                    const data = await res.json();
                    instr.id = data.id;
                    instr.isNew = false;
                    Toastify({ text: '✅ Инструкция добавлена', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                } else {
                    await fetch(`/api/instructions/${instr.id}`, {
                        method: 'PUT',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(payload)
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
                    Toastify({ text: `✅ Импорт: создано ${data.objects_created || 0}, обновлено ${data.objects_updated || 0}`, duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showImportModal = false;
                    await this.loadObjects();
                    await this.loadRoomCategories();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка импорта', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        }
    }
}).mount('#adminApp');