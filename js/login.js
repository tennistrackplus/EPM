/**
 * BigQuery Excel Connector - Login & Auth Controller
 * Handles SSO authentication, OAuth token persistence & session state
 */

import { Utils } from './utils.js';
import { Navigation } from './navigation.js';

export const Login = {
  currentUser: null,

  init: () => {
    Login.currentUser = Utils.storage.get('auth_user', null);
    Login.updateHeaderUI();
  },

  isAuthenticated: () => {
    return Login.currentUser !== null;
  },

  authenticate: async (email, password) => {
    // Enterprise Mock SSO Authenticator
    if (!email || !password) {
      throw new Error('Please fill in all mandatory credentials.');
    }

    // Simulate network authentication latency
    await new Promise(resolve => setTimeout(resolve, 800));

    const userProfile = {
      name: email.split('@')[0].toUpperCase(),
      email: email,
      tenant: 'Enterprise License',
      avatarInitials: email.substring(0, 2).toUpperCase(),
      token: 'jwt_mock_token_enterprise_' + Date.now()
    };

    Login.currentUser = userProfile;
    Utils.storage.set('auth_user', userProfile);
    Login.updateHeaderUI();

    return userProfile;
  },

  logout: () => {
    Login.currentUser = null;
    Utils.storage.remove('auth_user');
    Login.updateHeaderUI();
    Navigation.navigateTo('login');
    Utils.showToast('Logged out successfully', 'info');
  },

  updateHeaderUI: () => {
    const badge = Utils.$('#connection-status-badge');
    const avatar = Utils.$('.avatar-placeholder');

    if (Login.isAuthenticated()) {
      if (badge) {
        badge.className = 'status-indicator status-connected';
        badge.title = `Connected as ${Login.currentUser.email}`;
      }
      if (avatar) {
        avatar.textContent = Login.currentUser.avatarInitials;
      }
    } else {
      if (badge) {
        badge.className = 'status-indicator status-disconnected';
        badge.title = 'Disconnected';
      }
      if (avatar) {
        avatar.textContent = 'G';
      }
    }
  }
};