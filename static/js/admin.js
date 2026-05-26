const { createApp } = Vue;

createApp({
    data() {
        return {
            objects: [],
            categories: [],
            roomCategories: [],
            products: [],
            searchQuery: '',
            sortField: 'sort_priority',
            sortDir: 'asc',

            showEditModal: false,
            showAddModal: false,
            showImportModal: false,
            showInstructionsModal: false,
            showProductForm: false,
            showProductImportModal: false,
            editingProduct: null,
            productForm: { name: '', product_type: '', color: '' },
            productImportFile: null,

            currentSection: 'objects',

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

            // ===== Новые секции =====
            objectProperties: [],
            objectGroups: [],
            cleaningMethods: [],

            // Поиск для свойств и групп
            searchPropertyQuery: '',
            searchGroupQuery: '',

            showPropertyForm: false,
            showGroupForm: false,
            showMethodForm: false,
            showPropertyImportModal: false,
            showGroupImportModal: false,
            showMethodImportModal: false,

            editingProperty: null,
            editingGroup: null,
            editingMethod: null,

            propertyForm: {
                object_id: null,
                is_split: false,
                is_multi_method: false,
                has_support_maintenance: false,
                special_product_type: ''
            },
            groupForm: {
                group_name: '',
                object_id: null
            },
            methodForm: {
                method_name: '',
                sort_order: 99
            },

            propertyImportFile: null,
            groupImportFile: null,
            methodImportFile: null,
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
        },
        filteredObjectProperties() {
            let result = [...this.objectProperties];
            if (this.searchPropertyQuery) {
                const q = this.searchPropertyQuery.toLowerCase();
                result = result.filter(p =>
                    (p.object_name || '').toLowerCase().includes(q) ||
                    (p.normalized_name || '').toLowerCase().includes(q)
                );
            }
            return result;
        },
        filteredObjectGroups() {
            let result = [...this.objectGroups];
            if (this.searchGroupQuery) {
                const q = this.searchGroupQuery.toLowerCase();
                result = result.filter(g =>
                    (g.group_name || '').toLowerCase().includes(q) ||
                    (g.object_name || '').toLowerCase().includes(q) ||
                    (g.normalized_name || '').toLowerCase().includes(q)
                );
            }
            return result;
        }
    },
    mounted() {
        this.loadObjects();
        this.loadCategories();
        this.loadRoomCategories();
        this.loadProducts();
        this.loadObjectProperties();
        this.loadObjectGroups();
        this.loadCleaningMethods();
    },
    methods: {
        // ========== Загрузка данных ==========
        async loadObjects() {
            try {
                const res = await fetch('/api/objects');
                const data = await res.json();
                this.objects = data.objects || [];
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки объектов', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
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
        async loadObjectProperties() {
            try {
                const res = await fetch('/api/object-properties');
                const data = await res.json();
                this.objectProperties = data.object_properties || [];
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки свойств', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async loadObjectGroups() {
            try {
                const res = await fetch('/api/object-groups');
                const data = await res.json();
                this.objectGroups = data.object_groups || [];
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки групп', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async loadCleaningMethods() {
            try {
                const res = await fetch('/api/cleaning-methods');
                const data = await res.json();
                this.cleaningMethods = data.cleaning_methods || [];
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки методов', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        toggleSort(field) {
            if (this.sortField === field) {
                this.sortDir = this.sortDir === 'asc' ? 'desc' : 'asc';
            } else {
                this.sortField = field;
                this.sortDir = 'asc';
            }
        },

        // ========== Очистка БД ==========
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
                await this.loadObjectProperties();
                await this.loadObjectGroups();
                await this.loadCleaningMethods();
            } catch (e) {
                Toastify({ text: '❌ Ошибка очистки БД', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        // ========== Средства ==========
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
                    Toastify({ text: `✅ Импорт средств: создано ${data.created || 0}, обновлено ${data.updated || 0}`, duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showProductImportModal = false;
                    await this.loadProducts();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка импорта средств', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        // ========== Категории помещений ==========
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

        // ========== Объекты ==========
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

        // ========== Инструкции ==========
        async openInstructions(obj) {
            try {
                const res = await fetch(`/api/objects/${obj.id}/instructions`);
                const data = await res.json();
                this.currentInstructions = (data.instructions || []).map(i => ({
                    ...i,
                    isNew: false,
                    room_category_id: i.room_category_id || null,
                    maintenance_type: i.maintenance_type || 'основная',
                    surface_type: i.surface_type || null,
                    application_method: i.application_method || ''
                }));
                this.editObject = obj;
                this.showInstructionsModal = true;
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки инструкций', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        addInstruction() {
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
                    application_method: instr.application_method || '',
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

        // ========== Импорт объектов ==========
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
                    await this.loadObjectProperties();
                    await this.loadObjectGroups();
                    await this.loadCleaningMethods();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка импорта', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        // ================= СВОЙСТВА ОБЪЕКТОВ =================
        async updateProperty(prop) {
            try {
                await fetch(`/api/object-properties/${prop.id}`, {
                    method: 'PUT',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({
                        is_split: prop.is_split,
                        is_multi_method: prop.is_multi_method,
                        has_support_maintenance: prop.has_support_maintenance,
                        special_product_type: prop.special_product_type || ''
                    })
                });
                Toastify({ text: '✅ Свойство обновлено', duration: 2000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
            } catch (e) {
                Toastify({ text: '❌ Ошибка обновления свойства', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async deleteProperty(prop) {
            if (!confirm('Удалить свойство?')) return;
            try {
                await fetch(`/api/object-properties/${prop.id}`, { method: 'DELETE' });
                Toastify({ text: '✅ Свойство удалено', duration: 2000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                await this.loadObjectProperties();
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        openPropertyForm() {
            this.editingProperty = null;
            this.propertyForm = { object_id: null, is_split: false, is_multi_method: false, has_support_maintenance: false, special_product_type: '' };
            this.showPropertyForm = true;
        },
        async saveProperty() {
            const payload = {
                object_id: this.propertyForm.object_id,
                is_split: this.propertyForm.is_split,
                is_multi_method: this.propertyForm.is_multi_method,
                has_support_maintenance: this.propertyForm.has_support_maintenance,
                special_product_type: this.propertyForm.special_product_type
            };
            try {
                const url = this.editingProperty ? `/api/object-properties/${this.editingProperty.id}` : '/api/object-properties';
                const method = this.editingProperty ? 'PUT' : 'POST';
                const res = await fetch(url, {
                    method,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                if (res.ok) {
                    Toastify({ text: '✅ Свойство сохранено', duration: 2000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showPropertyForm = false;
                    await this.loadObjectProperties();
                } else {
                    const data = await res.json();
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        // Импорт/экспорт свойств
        handlePropertyImportFile(e) { this.propertyImportFile = e.target.files[0]; },
        async doPropertyImport() {
            if (!this.propertyImportFile) return;
            const formData = new FormData();
            formData.append('file', this.propertyImportFile);
            try {
                const res = await fetch('/api/object-properties/import_csv', { method: 'POST', body: formData });
                const data = await res.json();
                if (data.success) {
                    Toastify({ text: `✅ Импорт свойств: создано ${data.created || 0}, обновлено ${data.updated || 0}`, duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showPropertyImportModal = false;
                    await this.loadObjectProperties();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка импорта свойств', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        // ================= ГРУППЫ ОБЪЕКТОВ =================
        openGroupForm() {
            this.editingGroup = null;
            this.groupForm = { group_name: '', object_id: null };
            this.showGroupForm = true;
        },
        editGroup(group) {
            this.editingGroup = group;
            this.groupForm = { group_name: group.group_name, object_id: group.object_id };
            this.showGroupForm = true;
        },
        async saveGroup() {
            const payload = {
                group_name: this.groupForm.group_name.trim(),
                object_id: this.groupForm.object_id
            };
            if (!payload.group_name || !payload.object_id) {
                Toastify({ text: '⚠️ Заполните все поля', duration: 3000, gravity: 'bottom', position: 'right' }).showToast();
                return;
            }
            try {
                const url = this.editingGroup ? `/api/object-groups/${this.editingGroup.id}` : '/api/object-groups';
                const method = this.editingGroup ? 'PUT' : 'POST';
                const res = await fetch(url, {
                    method,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                if (res.ok) {
                    Toastify({ text: '✅ Группа сохранена', duration: 2000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showGroupForm = false;
                    await this.loadObjectGroups();
                } else {
                    const data = await res.json();
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async deleteGroup(group) {
            if (!confirm('Удалить группу?')) return;
            try {
                await fetch(`/api/object-groups/${group.id}`, { method: 'DELETE' });
                Toastify({ text: '✅ Группа удалена', duration: 2000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                await this.loadObjectGroups();
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        handleGroupImportFile(e) { this.groupImportFile = e.target.files[0]; },
        async doGroupImport() {
            if (!this.groupImportFile) return;
            const formData = new FormData();
            formData.append('file', this.groupImportFile);
            try {
                const res = await fetch('/api/object-groups/import_csv', { method: 'POST', body: formData });
                const data = await res.json();
                if (data.success) {
                    Toastify({ text: `✅ Импорт групп: создано ${data.created || 0}`, duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showGroupImportModal = false;
                    await this.loadObjectGroups();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка импорта групп', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },

        // ================= МЕТОДЫ ОЧИСТКИ =================
        openMethodForm() {
            this.editingMethod = null;
            this.methodForm = { method_name: '', sort_order: 99 };
            this.showMethodForm = true;
        },
        editMethod(m) {
            this.editingMethod = m;
            this.methodForm = { method_name: m.method_name, sort_order: m.sort_order };
            this.showMethodForm = true;
        },
        async saveMethod() {
            const payload = {
                method_name: this.methodForm.method_name.trim(),
                sort_order: parseInt(this.methodForm.sort_order) || 99
            };
            if (!payload.method_name) {
                Toastify({ text: '⚠️ Введите название метода', duration: 3000, gravity: 'bottom', position: 'right' }).showToast();
                return;
            }
            try {
                const url = this.editingMethod ? `/api/cleaning-methods/${this.editingMethod.id}` : '/api/cleaning-methods';
                const method = this.editingMethod ? 'PUT' : 'POST';
                const res = await fetch(url, {
                    method,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                if (res.ok) {
                    Toastify({ text: '✅ Метод сохранен', duration: 2000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showMethodForm = false;
                    await this.loadCleaningMethods();
                } else {
                    const data = await res.json();
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        async deleteMethod(m) {
            if (!confirm('Удалить метод?')) return;
            try {
                await fetch(`/api/cleaning-methods/${m.id}`, { method: 'DELETE' });
                Toastify({ text: '✅ Метод удален', duration: 2000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                await this.loadCleaningMethods();
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        handleMethodImportFile(e) { this.methodImportFile = e.target.files[0]; },
        async doMethodImport() {
            if (!this.methodImportFile) return;
            const formData = new FormData();
            formData.append('file', this.methodImportFile);
            try {
                const res = await fetch('/api/cleaning-methods/import_csv', { method: 'POST', body: formData });
                const data = await res.json();
                if (data.success) {
                    Toastify({ text: `✅ Импорт методов: создано ${data.created || 0}, обновлено ${data.updated || 0}`, duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    this.showMethodImportModal = false;
                    await this.loadCleaningMethods();
                } else {
                    Toastify({ text: '❌ ' + (data.error || 'Ошибка'), duration: 5000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка импорта методов', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        }
    }
}).mount('#adminApp');