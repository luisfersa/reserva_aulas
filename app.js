/**
 * App.js: Lógica final optimizada para Microsoft Teams
 * Filtros persistentes, multiselección y validación de multi-reserva.
 */

const App = (() => {
    const state = {
        aulas: [],
        bloques: [],
        calendario: [],
        reservas: [],
        user: null,
        currentStart: new Date(),
        // Filtros
        selectedAulas: [], // IDs de aulas seleccionadas
        showOccupied: false,
        isLoading: true,
        selectedSlot: null
    };

    const el = {
        container: document.getElementById('app-container'),
        rangeText: document.getElementById('current-range'),
        userInfo: document.getElementById('user-info'),
        modal: document.getElementById('booking-modal'),
        bookingForm: document.getElementById('booking-form-inner'),
        aulaSelector: document.getElementById('aula-selector'),
        toggleOccupied: document.getElementById('toggle-occupied')
    };

    const init = async () => {
        state.user = API.getUser();
        el.userInfo.textContent = `${state.user.nombre} ${state.user.isOwner ? '(Adm)' : ''}`;
        
        setStartToMonday(new Date());
        setupBaseEvents();
        await refreshAll();
        
        window.onresize = render;
    };

    const setStartToMonday = (d) => {
        const date = new Date(d);
        const day = date.getDay();
        const diff = date.getDate() - day + (day === 0 ? -6 : 1);
        state.currentStart = new Date(date.setDate(diff));
        state.currentStart.setHours(0,0,0,0);
    };

    const refreshAll = async () => {
        state.isLoading = true;
        renderLoader();
        try {
            const s = state.currentStart.toISOString().split('T')[0];
            const eD = new Date(state.currentStart); eD.setDate(eD.getDate() + 6);
            const e = eD.toISOString().split('T')[0];

            [state.aulas, state.bloques, state.calendario, state.reservas] = await Promise.all([
                API.fetchAulas(),
                API.fetchBloques(),
                API.fetchCalendario(),
                API.fetchReservas(s, e)
            ]);

            // Inicializar filtros si es la primera vez
            if (state.selectedAulas.length === 0) {
                state.selectedAulas = state.aulas.map(a => a.id);
                renderFilters();
            }
        } catch (err) { console.error(err); }
        finally {
            state.isLoading = false;
            render();
        }
    };

    const renderFilters = () => {
        el.aulaSelector.innerHTML = state.aulas.map(a => `
            <label>
                <input type="checkbox" value="${a.id}" checked onchange="App.toggleAula(${a.id}, this.checked)">
                <span>${a.nombre}</span>
            </label>
        `).join('');
    };

    const renderLoader = () => { el.container.innerHTML = `<div class="loader-container"><div class="loader"></div><p>Cargando cuadrícula...</p></div>`; };

    const render = () => {
        if (state.isLoading) return;
        updateNav();
        const isMobile = window.innerWidth < 1024;
        isMobile ? renderMobile() : renderDesktop();
    };

    const updateNav = () => {
        const end = new Date(state.currentStart); end.setDate(end.getDate() + 4);
        const opt = { day: 'numeric', month: 'short' };
        el.rangeText.textContent = `${state.currentStart.toLocaleDateString('es-ES', opt)} - ${end.toLocaleDateString('es-ES', opt)}`;
    };

    const getDayFormat = (date) => {
        const d = date.toLocaleDateString('es-ES', { weekday: 'short' }).toUpperCase().replace('.', '');
        const m = date.toLocaleDateString('es-ES', { month: 'short' }).toUpperCase().replace('.', '');
        return `${d} <span>${date.getDate()}</span> ${m}`;
    };

    const getHoliday = (date) => {
        const ds = date.toISOString().split('T')[0];
        const entry = state.calendario.find(c => ds >= c.fecha_inicio && ds <= (c.fecha_fin || c.fecha_inicio));
        return entry && !entry.es_lectivo ? entry : null;
    };

    const renderDesktop = () => {
        let h = `<div class="schedule-grid"><div class="grid-header">TRÁMITE</div>`;
        const days = getWeekDays();
        days.forEach(d => h += `<div class="grid-header">${getDayFormat(d)}</div>`);

        state.bloques.forEach(b => {
            h += `<div class="time-col"><strong>${b.nombre}</strong><small>${b.hora_inicio}-${b.hora_fin}</small></div>`;
            days.forEach(day => {
                const ds = day.toISOString().split('T')[0];
                const holiday = getHoliday(day);
                const past = day < new Date().setHours(0,0,0,0);

                if (holiday) { 
                    h += `<div class="slot holiday">
                            <span class="holiday-icon">🏖️</span>
                            <span class="holiday-desc">${holiday.descripcion || 'Festivo'}</span>
                          </div>`; 
                }
                else {
                    h += `<div class="slot">`;
                    state.aulas.filter(a => state.selectedAulas.includes(a.id)).forEach(aula => {
                        const res = state.reservas.find(r => r.fecha === ds && r.aula_id === aula.id && r.bloque_id === b.id);
                        const pcInfo = `<span class="pc-tag">${aula.num_pcs}🖥️</span>`;
                        if (res) {
                            const isOwn = res.usuario_email === state.user.email;
                            const cls = isOwn ? 'occupied-own' : (state.showOccupied ? 'occupied-subtle' : 'hidden');
                            const label = res.grupo_asignatura || res.usuario_nombre.split(' ')[0];
                            h += `<div class="aula-card ${cls}" onclick="App.openModal(${aula.id},'${ds}',${b.id},${res.id})">${aula.nombre}${pcInfo}: ${label}</div>`;
                        } else if (!past) {
                            h += `<div class="aula-card free" onclick="App.openModal(${aula.id},'${ds}',${b.id})">${aula.nombre} ${pcInfo}</div>`;
                        }
                    });
                    h += `</div>`;
                }
            });
        });
        h += `</div>`;
        el.container.innerHTML = h;
    };

    const renderMobile = () => {
        let h = '';
        const today = new Date().setHours(0,0,0,0);
        const days = getWeekDays().filter(d => d >= today);

        days.forEach(day => {
            const ds = day.toISOString().split('T')[0];
            const holiday = getHoliday(day);
            h += `<div class="mobile-day-card"><div class="day-title">${day.toLocaleDateString('es-ES', { weekday: 'long', day: 'numeric', month: 'long' }).toUpperCase()}</div>`;
            if (holiday) { 
                h += `<div class="holiday-msg">
                        <span class="msg-icon">🏖️</span>
                        <p>${holiday.descripcion || 'Festivo'}</p>
                      </div>`; 
            }
            else {
                h += `<div class="availability-matrix">`;
                state.bloques.forEach(b => {
                    h += `<div class="matrix-row"><div class="row-label">${b.nombre}</div>`;
                    state.aulas.filter(a => state.selectedAulas.includes(a.id)).forEach(aula => {
                        const res = state.reservas.find(r => r.fecha === ds && r.aula_id === aula.id && r.bloque_id === b.id);
                        const pcTag = `<small style="font-size:0.5rem; display:block">${aula.num_pcs}P</small>`;
                        if (res) {
                            const isOwn = res.usuario_email === state.user.email;
                            const cls = isOwn ? 'occupied-own' : (state.showOccupied ? 'occupied-subtle' : 'hidden');
                            const label = res.grupo_asignatura ? res.grupo_asignatura.substring(0,5) : res.usuario_nombre.substring(0,3);
                            h += `<div class="matrix-cell ${cls}" onclick="App.openModal(${aula.id},'${ds}',${b.id},${res.id})"><div>${label}${pcTag}</div></div>`;
                        } else {
                            h += `<div class="matrix-cell free" onclick="App.openModal(${aula.id},'${ds}',${b.id})"><div>${aula.nombre.substring(0,3)}${pcTag}</div></div>`;
                        }
                    });
                    h += `</div>`;
                });
                h += `</div>`;
            }
            h += `</div>`;
        });
        el.container.innerHTML = h || '<p>No hay días lectivos.</p>';
    };

    const getWeekDays = () => {
        const res = [];
        for(let i=0; i<5; i++) {
            const d = new Date(state.currentStart); d.setDate(d.getDate() + i);
            res.push(d);
        }
        return res;
    };

    // --- ACCIONES ---

    const toggleAula = (id, checked) => {
        if (checked) state.selectedAulas.push(id);
        else state.selectedAulas = state.selectedAulas.filter(x => x !== id);
        render();
    };

    const openModal = (aulaId, fecha, bloqueId, reservaId = null) => {
        const aula = state.aulas.find(a => a.id === aulaId);
        const bloque = state.bloques.find(b => b.id === bloqueId);
        const res = reservaId ? state.reservas.find(r => r.id === reservaId) : null;
        
        state.selectedSlot = { aulaId, fecha, bloqueId, reservaId };
        
        const fFmt = new Date(fecha).toLocaleDateString('es-ES', { dateStyle: 'long' });
        document.getElementById('modal-title').textContent = res ? 'Detalle de Reserva' : 'Nueva Reserva';
        document.getElementById('details-text').textContent = `${aula.nombre} | ${fFmt} | ${bloque.nombre}`;
        document.getElementById('form-pcs').textContent = aula.num_pcs;
        document.getElementById('form-obs').textContent = aula.observaciones || '---';
        document.getElementById('form-grupo-asignatura').value = res ? res.grupo_asignatura : '';
        
        const bSave = document.getElementById('btn-save');
        const bDel = document.getElementById('btn-cancel-reservation');

        if (res) {
            bSave.classList.add('hidden');
            bDel.classList.toggle('hidden', !(state.user.isOwner || res.usuario_email === state.user.email));
        } else {
            bSave.classList.remove('hidden');
            bDel.classList.add('hidden');
        }
        el.modal.style.display = 'block';
    };

    const setupBaseEvents = () => {
        document.getElementById('prev-week').onclick = () => { state.currentStart.setDate(state.currentStart.getDate() - 7); refreshAll(); };
        document.getElementById('next-week').onclick = () => { state.currentStart.setDate(state.currentStart.getDate() + 7); refreshAll(); };
        document.querySelector('.close-modal').onclick = () => el.modal.style.display = 'none';
        
        el.toggleOccupied.onchange = (e) => { state.showOccupied = e.target.checked; render(); };

        el.bookingForm.onsubmit = async (e) => {
            e.preventDefault();
            const { aulaId, fecha, bloqueId } = state.selectedSlot;
            
            // Validación Multi-reserva
            if (!state.user.isOwner) {
                const exist = state.reservas.find(r => r.fecha === fecha && r.bloque_id === bloqueId && r.usuario_email === state.user.email);
                if (exist) {
                    alert("Lo sentimos, ya tienes otra aula reservada en este mismo tramo hoy.");
                    return;
                }
            }

            try {
                await API.saveReserva({
                    aula_id: aulaId, fecha, bloque_id: bloqueId,
                    usuario_email: state.user.email,
                    usuario_nombre: state.user.nombre,
                    grupo_asignatura: document.getElementById('form-grupo-asignatura').value || ''
                });
                el.modal.style.display = 'none';
                await refreshAll();
            } catch(e) { alert("Error al guardar."); }
        };

        document.getElementById('btn-cancel-reservation').onclick = async () => {
            if (!confirm("¿Anular reserva?")) return;
            try {
                await API.deleteReserva(state.selectedSlot.reservaId);
                el.modal.style.display = 'none';
                await refreshAll();
            } catch(e) { alert("Error al borrar."); }
        };

        window.onclick = (ev) => { if (ev.target === el.modal) el.modal.style.display = 'none'; };
    };

    return { init, toggleAula, openModal };
})();

document.addEventListener('DOMContentLoaded', App.init);
