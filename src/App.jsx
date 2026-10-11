import { useEffect, useState } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';

const tenantId = 'YOUR_TENANT_ID';
const databricksWorkspaceUrl = 'https://<your-workspace>.cloud.databricks.com';
const databricksResourceAppId = 'YOUR_DATABRICKS_APP_ID';
const genieMcpUrl = 'https://<your-workspace>.cloud.databricks.com/api/2.0/mcp/genie/<your-genie-id>';
const genieQuestion = 'What is the distribution of reviews by franchiseID?';

const msalConfig = {
  auth: {
    clientId: 'YOUR_CLIENT_ID',
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false
  }
};

const graphScopes = ['User.Read'];
const databricksScopes = [`api://${databricksResourceAppId}/user_impersonation`];
const msalInstance = new PublicClientApplication(msalConfig);

const initializeMsal = async () => {
  await msalInstance.initialize();
};

const validateTenant = (account) => {
  if (!account || !account.tenantId) {
    return false;
  }

  return account.tenantId === tenantId || tenantId === 'YOUR_TENANT_ID';
};

const getGraphProfile = async (account) => {
  const silentResponse = await msalInstance.acquireTokenSilent({
    account,
    scopes: graphScopes
  });

  const response = await fetch('https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName', {
    headers: {
      Authorization: `Bearer ${silentResponse.accessToken}`
    }
  });

  if (!response.ok) {
    throw new Error(`Graph request failed: ${response.status}`);
  }

  return response.json();
};

const getDatabricksAccessTokenForSignedInUser = async () => {
  const accounts = msalInstance.getAllAccounts();
  const account = accounts[0];

  if (!account) {
    throw new Error('No signed-in Azure account available for Databricks.');
  }

  if (!validateTenant(account)) {
    throw new Error('Signed-in user is not part of the configured Azure tenant.');
  }

  if (!databricksResourceAppId || databricksResourceAppId.trim() === '' || databricksResourceAppId === 'YOUR_DATABRICKS_APP_ID') {
    throw new Error('Set the real Databricks app/resource ID in App.jsx before calling Databricks.');
  }

  try {
    const silentResponse = await msalInstance.acquireTokenSilent({
      account,
      scopes: databricksScopes
    });

    return silentResponse.accessToken;
  } catch (silentError) {
    console.warn('Databricks silent token request failed. Falling back to interactive popup.', silentError);

    const popupResponse = await msalInstance.acquireTokenPopup({
      account,
      scopes: databricksScopes,
      prompt: 'select_account'
    });

    return popupResponse.accessToken;
  }
};

const runDatabricksGenie = async () => {
  const accessToken = await getDatabricksAccessTokenForSignedInUser();

  const response = await fetch(genieMcpUrl, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      question: genieQuestion,
      prompt: genieQuestion,
      messages: [
        {
          role: 'user',
          content: genieQuestion
        }
      ]
    })
  });

  if (!response.ok) {
    throw new Error(`Databricks Genie request failed: ${response.status}`);
  }

  return response.json();
};

function App() {
  const [status, setStatus] = useState('Checking Entra session...');
  const [statusType, setStatusType] = useState('info');
  const [signedIn, setSignedIn] = useState(false);
  const [profile, setProfile] = useState(null);
  const [databricksProfile, setDatabricksProfile] = useState(null);
  const [databricksLoading, setDatabricksLoading] = useState(false);
  const [genieResponse, setGenieResponse] = useState(null);
  const [genieLoading, setGenieLoading] = useState(false);

  const setMessage = (message, type = 'info') => {
    setStatus(message);
    setStatusType(type);
  };

  useEffect(() => {
    const initialize = async () => {
      try {
        await initializeMsal();
      } catch (error) {
        console.error('MSAL initialize failed', error);
        setMessage('MSAL failed to initialize.', 'warning');
        return;
      }

      const accounts = msalInstance.getAllAccounts();

      if (accounts.length === 0) {
        setSignedIn(false);
        setProfile(null);
        setDatabricksProfile(null);
        setMessage('No active Microsoft Entra session found. Sign in to continue.', 'info');
        return;
      }

      const currentAccount = accounts[0];
      if (!validateTenant(currentAccount)) {
        setSignedIn(false);
        setProfile(null);
        setDatabricksProfile(null);
        setMessage('This user is not part of the configured tenant.', 'warning');
        return;
      }

      try {
        const profileData = await getGraphProfile(currentAccount);
        setProfile(profileData);
        setSignedIn(true);
        setMessage('Signed in successfully using Microsoft Entra ID.', 'success');
      } catch (error) {
        console.error(error);
        setSignedIn(false);
        setProfile(null);
        setDatabricksProfile(null);
        setMessage('Unable to load the signed-in profile. Please sign in again.', 'warning');
      }
    };

    initialize();
  }, []);

  const signIn = async () => {
    try {
      await initializeMsal();
    } catch (error) {
      console.error('MSAL initialize failed before login', error);
      setMessage('MSAL failed to initialize.', 'warning');
      return;
    }

    const accounts = msalInstance.getAllAccounts();
    const existingAccount = accounts[0];

    if (existingAccount) {
      if (!validateTenant(existingAccount)) {
        setSignedIn(false);
        setProfile(null);
        setDatabricksProfile(null);
        setMessage('This user is not part of the configured tenant.', 'warning');
        return;
      }

      try {
        setMessage('Using the current Microsoft Entra session...', 'info');
        const profileData = await getGraphProfile(existingAccount);
        setProfile(profileData);
        setSignedIn(true);
        setMessage('Signed in successfully using your current session.', 'success');
        return;
      } catch (error) {
        console.warn('Silent session failed. Opening Microsoft sign-in with account hint.', error);
      }
    }

    try {
      setMessage('Opening Microsoft sign-in popup...', 'info');
      const loginResponse = await msalInstance.loginPopup({
        scopes: graphScopes,
        prompt: 'select_account',
        loginHint: existingAccount ? existingAccount.username : undefined
      });

      if (!validateTenant(loginResponse.account)) {
        setSignedIn(false);
        setProfile(null);
        setDatabricksProfile(null);
        setMessage('This user is not part of the configured tenant.', 'warning');
        return;
      }

      const profileData = await getGraphProfile(loginResponse.account);
      setProfile(profileData);
      setSignedIn(true);
      setMessage('Signed in successfully using Microsoft Entra ID.', 'success');
    } catch (error) {
      console.error(error);
      setSignedIn(false);
      setMessage('Sign-in was cancelled, failed, or the user is not part of this tenant.', 'warning');
    }
  };

  const signOut = async () => {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      try {
        await msalInstance.logoutPopup({ account: accounts[0] });
        return;
      } catch (error) {
        console.error(error);
      }
    }

    setSignedIn(false);
    setProfile(null);
    setDatabricksProfile(null);
    setGenieResponse(null);
    setMessage('Signed out. No active Entra user session remains.', 'info');
  };

  const handleOpenDatabricks = () => {
    if (!databricksWorkspaceUrl || databricksWorkspaceUrl.includes('<your-workspace>')) {
      setMessage('Set your Databricks workspace URL in App.jsx before opening Databricks.', 'warning');
      return;
    }

    window.open(databricksWorkspaceUrl, '_blank', 'noopener,noreferrer');
  };

  const handleDatabricksProfile = async () => {
    setDatabricksLoading(true);
    setMessage('Loading Databricks profile...', 'info');

    try {
      const accessToken = await getDatabricksAccessTokenForSignedInUser();
      const response = await fetch(`${databricksWorkspaceUrl}/api/2.0/preview/scim/v2/Me`, {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      });

      if (!response.ok) {
        throw new Error(`Databricks profile request failed: ${response.status}`);
      }

      const databricksUser = await response.json();
      setDatabricksProfile(databricksUser);
      setMessage('Databricks profile loaded successfully.', 'success');
    } catch (error) {
      console.error(error);
      setDatabricksProfile(null);
      setMessage('Unable to load Databricks profile. Check workspace URL, resource app ID, and Databricks access.', 'warning');
    } finally {
      setDatabricksLoading(false);
    }
  };

  const handleGenieQuestion = async () => {
    setGenieLoading(true);
    setMessage('Requesting Databricks access token...', 'info');

    try {
      const accessToken = await getDatabricksAccessTokenForSignedInUser();
      setGenieResponse({
        tokenAcquired: true,
        tokenPreview: `${accessToken.substring(0, 20)}...`,
        message: 'Databricks access token was returned successfully.'
      });
      setMessage('Databricks token acquired successfully. The Genie question is not being executed yet.', 'success');
    } catch (error) {
      console.error(error);
      setGenieResponse(null);
      setMessage('Databricks token request failed. Check Databricks access and app setup.', 'warning');
    } finally {
      setGenieLoading(false);
    }
  };

  const displayName = profile?.displayName || profile?.userPrincipalName || 'Unknown user';
  const email = profile?.userPrincipalName || 'No email available';
  const objectId = profile?.id || 'Not available';
  const initials = (displayName || 'U').charAt(0).toUpperCase();

  return (
    <>
      <div className="background-glow glow-one"></div>
      <div className="background-glow glow-two"></div>

      <main className="container">
        <section className="login-shell">
          <div className="brand-panel">
            <div className="brand-logo">M</div>
            <div>
              <p className="eyebrow">Secure access</p>
              <h1>Microsoft Entra SSO</h1>
            </div>
            <p className="brand-copy">
              Sign in with your Microsoft account and use the same identity to access Databricks.
            </p>
          </div>

          <div className="form-panel">
            <div className="header-row">
              <div>
                <p className="eyebrow accent">Workspace</p>
                <h2>Welcome</h2>
              </div>
              <span className="status-badge">SSO Test</span>
            </div>

            <p className={`status ${statusType}`}>{status}</p>

            {!signedIn ? (
              <div className="view">
                <p className="helper-text">No active Microsoft Entra user session was found.</p>
                <button className="primary-btn" onClick={signIn}>
                  <span className="btn-icon">🪟</span>
                  Sign in with Microsoft
                </button>
              </div>
            ) : (
              <div className="view">
                <div className="profile-header">
                  <div className="avatar">{initials}</div>
                  <div>
                    <h3>{displayName}</h3>
                    <p className="user-email">{email}</p>
                  </div>
                </div>

                <div className="details">
                  <div className="detail-row">
                    <span className="label">User Principal Name</span>
                    <span>{profile?.userPrincipalName || '-'}</span>
                  </div>
                  <div className="detail-row">
                    <span className="label">Object ID</span>
                    <span>{objectId}</span>
                  </div>
                  <div className="detail-row">
                    <span className="label">Azure Profile</span>
                    <span>{profile?.displayName || '-'}</span>
                  </div>
                </div>

                <div className="databricks-actions" style={{ marginTop: '18px', display: 'grid', gap: '12px' }}>
                  <button className="primary-btn" onClick={handleOpenDatabricks}>
                    Open Databricks Workspace
                  </button>

                  <button className="secondary-btn" onClick={handleDatabricksProfile} disabled={databricksLoading}>
                    {databricksLoading ? 'Loading Databricks profile...' : 'Load Databricks profile'}
                  </button>

                  <button className="secondary-btn" onClick={handleGenieQuestion} disabled={genieLoading}>
                    {genieLoading ? 'Running Genie...' : 'Ask Genie: What is the distribution of reviews by franchiseID?'}
                  </button>
                </div>

                {databricksProfile && (
                  <div className="details" style={{ marginTop: '18px' }}>
                    <div className="detail-row">
                      <span className="label">Databricks Display Name</span>
                      <span>{databricksProfile?.displayName || '-'}</span>
                    </div>
                    <div className="detail-row">
                      <span className="label">Databricks User Name</span>
                      <span>{databricksProfile?.userName || '-'}</span>
                    </div>
                    <div className="detail-row">
                      <span className="label">Databricks Groups</span>
                      <span>{databricksProfile?.groups?.map((g) => g.display).join(', ') || '-'}</span>
                    </div>
                  </div>
                )}

                {genieResponse && (
                  <div className="details" style={{ marginTop: '18px' }}>
                    <div className="detail-row">
                      <span className="label">Genie Question</span>
                      <span>{genieQuestion}</span>
                    </div>
                    <div className="detail-row">
                      <span className="label">Genie Result</span>
                      <span>{JSON.stringify(genieResponse, null, 2)}</span>
                    </div>
                  </div>
                )}

                <button className="secondary-btn" onClick={signOut}>Sign out</button>
              </div>
            )}
          </div>
        </section>
      </main>
    </>
  );
}

export default App;
