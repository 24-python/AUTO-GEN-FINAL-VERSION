const { createApp } = Vue;

createApp({
    data() {
        return {
            history: [],
            page: 1,
            totalPages: 1,
            total: 0,
            showDateDelete: false,
            deleteBeforeDate: ''
        }
    },
    computed: {
        allSelected() {
            return this.history.length > 0 && this.history.every(h => h.selected);
        },
        selectedCount() {
            return this.history.filter(h => h.selected).length;
        },
        displayPages() {
            const pages = [];
            const start = Math.max(1, this.page - 2);
            const end = Math.min(this.totalPages, this.page + 2);
            for (let i = start; i <= end; i++) pages.push(i);
            return pages;
        }
    },
    mounted() {
        this.loadHistory();
    },
    methods: {
        async loadHistory() {
            try {
                const res = await fetch(`/api/history?page=${this.page}`);
                const data = await res.json();
                this.history = (data.history || []).map(h => ({...h, selected: false}));
                this.totalPages = data.total_pages;
                this.total = data.total;
            } catch (e) {
                Toastify({ text: '❌ Ошибка загрузки истории', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        },
        toggleAll(e) {
            const checked = e.target.checked;
            this.history.forEach(h => h.selected = checked);
        },
        selectAll() {
            this.history.forEach(h => h.selected = true);
        },
        async deleteSelected() {
            const filenames = this.history.filter(h => h.selected).map(h => h.filename);
            if (filenames.length === 0) return;
            if (!confirm(`Удалить ${filenames.length} техкарт(ы)?`)) return;
            await this.deleteHistory(filenames);
        },
        async deleteByDate() {
            if (!this.deleteBeforeDate) return;
            if (!confirm(`Удалить техкарты старше ${this.deleteBeforeDate}?`)) return;
            await this.deleteHistory([], this.deleteBeforeDate);
        },
        async deleteHistory(filenames = [], beforeDate = null) {
            try {
                const res = await fetch('/api/history/delete', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({filenames, before_date: beforeDate})
                });
                const data = await res.json();
                if (data.success) {
                    Toastify({ text: `✅ Удалено: ${data.deleted} файлов`, duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#00A650' } }).showToast();
                    await this.loadHistory();
                }
            } catch (e) {
                Toastify({ text: '❌ Ошибка удаления', duration: 3000, gravity: 'bottom', position: 'right', style: { background: '#DC2626' } }).showToast();
            }
        }
    },
    watch: {
        page() {
            this.loadHistory();
        }
    }
}).mount('#historyApp');