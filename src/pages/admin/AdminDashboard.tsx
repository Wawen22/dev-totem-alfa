import React, { useState, useEffect, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import { useIsAdmin } from '../../hooks/useIsAdmin';
import { useAuthenticatedGraphClient } from '../../hooks/useAuthenticatedGraphClient';
import { SharePointService } from '../../services/sharePointService';
import { SharePointListItem } from '../../types/sharepoint';
import { forgiatiColumns } from '../../config/forgiatiColumns';
import { tubiColumns } from '../../config/tubiColumns';

export const AdminDashboard: React.FC = () => {
    const isAdmin = useIsAdmin();
    const navigate = useNavigate();
    const [activeTab, setActiveTab] = useState<'forgiati' | 'tubi'>('forgiati');
    const [items, setItems] = useState<SharePointListItem[]>([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);

    const siteId = import.meta.env.VITE_SHAREPOINT_SITE_ID;
    const forgiatiListId = import.meta.env.VITE_FORGIATI_LIST_ID;
    const tubiListId = import.meta.env.VITE_TUBI_LIST_ID;
    const getClient = useAuthenticatedGraphClient();

    const sharePointService = useMemo(() => {
        if (!siteId) return null;
        return new SharePointService(getClient, siteId);
    }, [getClient, siteId]);

    useEffect(() => {
        if (!isAdmin) {
            navigate('/');
            return;
        }

        const fetchItems = async () => {
            if (!sharePointService) return;
            
            const listId = activeTab === 'forgiati' ? forgiatiListId : tubiListId;
            if (!listId) {
                setError(`List ID for ${activeTab} not configured in .env`);
                return;
            }

            setLoading(true);
            setError(null);
            try {
                const data = await sharePointService.listItems(listId);
                setItems(data);
            } catch (err: any) {
                console.error(err);
                setError(err.message || "Errore durante il caricamento");
            } finally {
                setLoading(false);
            }
        };

        fetchItems();
    }, [isAdmin, navigate, activeTab, sharePointService, forgiatiListId, tubiListId]);

    const activeColumns = useMemo(() => {
        return activeTab === 'forgiati' ? forgiatiColumns : tubiColumns;
    }, [activeTab]);

    if (!isAdmin) return null;

    return (
        <div className="totem-shell">
            <div className="hero">
                <div className="hero-content">
                    <button className="back-btn" onClick={() => navigate('/')}>
                        ←
                    </button>
                    <div>
                        <h1>Pannello Amministratore</h1>
                        <p className="eyebrow">GESTIONE LOTTI E ARTICOLI</p>
                    </div>
                </div>
            </div>

            <div className="admin-main">
                <div className="admin-tabs">
                    <button 
                        className={`admin-tab ${activeTab === 'forgiati' ? 'active' : ''}`}
                        onClick={() => setActiveTab('forgiati')}
                    >
                        Forgiati
                    </button>
                    <button 
                        className={`admin-tab ${activeTab === 'tubi' ? 'active' : ''}`}
                        onClick={() => setActiveTab('tubi')}
                    >
                        Tubi
                    </button>
                </div>

                <div className="admin-grid">
                    <div className="panel">
                        <div className="panel-content">
                            <div className="section-heading">
                                <h3>Gestione {activeTab === 'forgiati' ? 'Forgiati' : 'Tubi'}</h3>
                                <button className="btn primary">
                                    + Aggiungi Articolo
                                </button>
                            </div>
                            
                            {error && <div className="alert error">{error}</div>}
                            
                            <div className="inventory-panel" style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
                                <div className="table-scroll">
                                    <table className="inventory-table">
                                        <thead>
                                            <tr>
                                                <th className="sticky-col">Codice / Title</th>
                                                {activeColumns.map(col => (
                                                    <th key={col.field}>{col.label}</th>
                                                ))}
                                                <th>Azioni</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {loading ? (
                                                <tr><td colSpan={activeColumns.length + 2} style={{textAlign: 'center', padding: 20}}>Caricamento in corso...</td></tr>
                                            ) : items.length === 0 ? (
                                                <tr><td colSpan={activeColumns.length + 2} style={{textAlign: 'center', padding: 20}}>Nessun elemento trovato.</td></tr>
                                            ) : (
                                                items.map(item => (
                                                    <tr key={item.id}>
                                                        <th className="sticky-col">
                                                            {/* Title is sometimes inside fields, sometimes not. Type def says fields & { Title?: string } */}
                                                            {item.fields.Title}
                                                        </th>
                                                        {activeColumns.map(col => (
                                                            <td key={col.field}>
                                                                {/* Just render string representation for now */}
                                                                {String(item.fields[col.field] || '')}
                                                            </td>
                                                        ))}
                                                        <td>
                                                            <button className="btn secondary" style={{padding: '4px 8px', fontSize: 13}}>
                                                                Modifica
                                                            </button>
                                                        </td>
                                                    </tr>
                                                ))
                                            )}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};
