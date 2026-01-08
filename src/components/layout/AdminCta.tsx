import React from 'react';
import { useNavigate } from 'react-router-dom';
import { useIsAdmin } from '../../hooks/useIsAdmin';

export const AdminCta: React.FC = () => {
    const isAdmin = useIsAdmin();
    const navigate = useNavigate();

    // If not admin, do not render anything
    if (!isAdmin) return null;

    return (
        <button 
            className="admin-cta"
            onClick={() => navigate('/admin')}
            aria-label="Admin Dashboard"
        >
            Admin Panel
        </button>
    );
};
