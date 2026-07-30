/**
 * BigQuery Excel Connector - Login & SSO Page
 */

import { Login } from '../login.js';
import { Navigation } from '../navigation.js';
import { Utils } from '../utils.js';

export async function renderLoginPage() {
  setTimeout(() => {
    const form = document.getElementById('login-form');
    if (form) {
      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const email = document.getElementById('login-email').value;
        const pass = document.getElementById('login-password').value;

        try {
          await Login.authenticate(email, pass);
          Utils.showToast('Inicio de sesión correcto', 'success');
          Navigation.navigateTo('home');
        } catch (err) {
          Utils.showToast(err.message, 'error');
        }
      });
    }
  }, 0);

  return `
    <div class="page-container p-base fade-in flex flex-col justify-center" style="min-height: 100%;">
      <div class="card p-base">
        <div class="text-center mb-base">
          <h1 class="text-lg font-bold">Enterprise Single Sign-On</h1>
          <p class="text-xs color-subtext">Accede con tus credenciales de Google Cloud o Directorio</p>
        </div>

        <form id="login-form" class="flex flex-col gap-sm">
          <div>
            <label class="text-xs font-semibold block mb-xs">Correo Electrónico</label>
            <input type="email" id="login-email" class="input-text w-full" placeholder="usuario@empresa.com" required value="analista@enterprise.com">
          </div>

          <div>
            <label class="text-xs font-semibold block mb-xs">Contraseña</label>
            <input type="password" id="login-password" class="input-text w-full" placeholder="••••••••" required value="password123">
          </div>

          <div class="mb-xs">
            <label class="text-xs font-semibold block mb-xs">Entorno GCP</label>
            <select class="input-text w-full" id="login-env">
              <option value="prod">Google Cloud Production</option>
              <option value="staging">Google Cloud Staging / Dev</option>
            </select>
          </div>

          <button type="submit" class="btn btn-primary w-full mt-xs">
            Iniciar Sesión SSO
          </button>
        </form>
      </div>
    </div>
  `;
}