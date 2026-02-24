import React, { useState, useEffect } from 'react';
import Login from './components/Login';
import Dashboard from './components/Dashboard';
import Setup from './components/Setup';
import db from './services/db';

function App() {
    const [isLoggedIn, setIsLoggedIn] = useState(false);
    const [userRole, setUserRole] = useState(null);
    const [isLoading, setIsLoading] = useState(true);

    const [needsSetup, setNeedsSetup] = useState(false);

    useEffect(() => {
        // Check if user is already logged in (persistent session)
        // and if the database needs initial setup
        const initApp = async () => {
            try {
                const isSetupRequired = await db.checkNeedsSetup();
                if (isSetupRequired) {
                    setNeedsSetup(true);
                    setIsLoading(false);
                    return;
                }

                const role = await db.getCurrentUserRole();
                if (role) {
                    setUserRole(role);
                    setIsLoggedIn(true);
                }
            } catch (error) {
                console.error('Init failed:', error);
            } finally {
                setIsLoading(false);
            }
        };
        initApp();
    }, []);

    const handleLogin = (role) => {
        setUserRole(role);
        setIsLoggedIn(true);
    };

    const handleLogout = async () => {
        await db.logout();
        setUserRole(null);
        setIsLoggedIn(false);
    };

    const handleSetupComplete = (role) => {
        setNeedsSetup(false);
        setUserRole(role);
        setIsLoggedIn(true);
    };

    if (isLoading) {
        return (
            <div className="loading-screen">
                <div className="loader"></div>
                <p>Loading...</p>
            </div>
        );
    }

    if (needsSetup) {
        return <Setup onSetupComplete={handleSetupComplete} />;
    }

    return (
        <div className="app">
            {isLoggedIn ? (
                <Dashboard userRole={userRole} onLogout={handleLogout} />
            ) : (
                <Login onLogin={handleLogin} />
            )}
        </div>
    );
}

export default App;
