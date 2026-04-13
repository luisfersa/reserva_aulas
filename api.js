/**
 * API final optimizada para listas SharePoint (snake_case)
 * Soporte para Microsoft Teams y normalización de datos.
 */

const API = (() => {
    const isSP = typeof _spPageContextInfo !== 'undefined';
    const siteUrl = isSP ? _spPageContextInfo.webAbsoluteUrl : '';
    
    const LISTS = {
        aulas: 'aulas',
        calendario: 'calendario_curso',
        bloques: 'bloques_horarios',
        reservas: 'reservas'
    };

    // Normalizador de horas (8:30 -> 08:30)
    const fixTime = (t) => {
        if (!t) return "00:00";
        const parts = t.split(':');
        return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}`;
    };

    // MOCKS para Teams Local Dev
    const mock = {
        aulas: [
            { id: 1, nombre: 'Aula 101', num_pcs: 25, observaciones: 'Planta 1' },
            { id: 2, nombre: 'Inf 1', num_pcs: 30, observaciones: 'Planta Baja' },
            { id: 3, nombre: 'Biblioteca', num_pcs: 12, observaciones: 'Silencio' }
        ],
        bloques: [
            { id: 1, nombre: '1ª', hora_inicio: '08:30', hora_fin: '09:25', orden: 1 },
            { id: 2, nombre: '2ª', hora_inicio: '09:25', hora_fin: '10:20', orden: 2 },
            { id: 3, nombre: 'Rec', hora_inicio: '11:15', hora_fin: '11:45', orden: 3 }
        ],
        calendario: [
            { fecha_inicio: '2024-12-20', fecha_fin: '2025-01-07', es_lectivo: false }
        ],
        reservas: []
    };

    const getDigest = async () => {
        if (!isSP) return "mock";
        const r = await fetch(`${siteUrl}/_api/contextinfo`, { method: 'POST', headers: { 'Accept': 'application/json;odata=verbose' } });
        const d = await r.json();
        return d.d.GetContextWebInformation.FormDigestValue;
    };

    return {
        getUser: () => {
            if (!isSP) return { email: 'profe@educa.jcyl.es', nombre: 'Profe Demo', isOwner: true };
            return {
                email: _spPageContextInfo.userEmail,
                nombre: _spPageContextInfo.userDisplayName,
                isOwner: _spPageContextInfo.isSiteAdmin || false
            };
        },

        fetchAulas: async () => {
            if (!isSP) return mock.aulas;
            const r = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${LISTS.aulas}')/items?$select=Id,nombre,num_pcs,observaciones`, {
                headers: { 'Accept': 'application/json;odata=nometadata' }
            });
            const d = await r.json();
            return d.value.map(i => ({ id: i.Id, nombre: i.nombre, num_pcs: i.num_pcs || 0, observaciones: i.observaciones || '' }));
        },

        fetchBloques: async () => {
            if (!isSP) return mock.bloques.sort((a,b) => a.orden - b.orden);
            const r = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${LISTS.bloques}')/items?$select=Id,nombre,hora_inicio,hora_fin,orden&$orderby=orden asc`, {
                headers: { 'Accept': 'application/json;odata=nometadata' }
            });
            const d = await r.json();
            return d.value.map(i => ({ id: i.Id, nombre: i.nombre, hora_inicio: fixTime(i.hora_inicio), hora_fin: fixTime(i.hora_fin), orden: i.orden }));
        },

        fetchCalendario: async () => {
            if (!isSP) return mock.calendario;
            const r = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${LISTS.calendario}')/items?$select=fecha_inicio,fecha_fin,es_lectivo,descripcion`, {
                headers: { 'Accept': 'application/json;odata=nometadata' }
            });
            const d = await r.json();
            return d.value.map(i => ({ 
                fecha_inicio: i.fecha_inicio.split('T')[0], 
                fecha_fin: i.fecha_fin ? i.fecha_fin.split('T')[0] : null, 
                es_lectivo: !!i.es_lectivo,
                descripcion: i.descripcion || ''
            }));
        },

        fetchReservas: async (s, e) => {
            if (!isSP) return mock.reservas;
            const filter = `fecha ge '${s}' and fecha le '${e}'`;
            const r = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${LISTS.reservas}')/items?$select=Id,aula_id,fecha,bloque_id,usuario_email,usuario_nombre,grupo_asignatura&$filter=${filter}`, {
                headers: { 'Accept': 'application/json;odata=nometadata' }
            });
            const d = await r.json();
            return d.value.map(i => ({
                id: i.Id, aula_id: i.aula_id, fecha: i.fecha.split('T')[0],
                bloque_id: i.bloque_id, usuario_email: i.usuario_email,
                usuario_nombre: i.usuario_nombre, grupo_asignatura: i.grupo_asignatura || ''
            }));
        },

        saveReserva: async (data) => {
            if (!isSP) {
                const nr = { id: Date.now(), ...data };
                mock.reservas.push(nr);
                return nr;
            }
            const digest = await getDigest();
            const r = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${LISTS.reservas}')/items`, {
                method: 'POST',
                body: JSON.stringify({ '__metadata': { 'type': `SP.Data.ReservasListItem` }, ...data }),
                headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest }
            });
            return await r.json();
        },

        deleteReserva: async (id) => {
            if (!isSP) {
                mock.reservas = mock.reservas.filter(r => r.id !== id);
                return true;
            }
            const digest = await getDigest();
            await fetch(`${siteUrl}/_api/web/lists/getbytitle('${LISTS.reservas}')/items(${id})`, {
                method: 'POST',
                headers: { 'X-RequestDigest': digest, 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }
            });
            return true;
        }
    };
})();
