# 🚀 Guía de Despliegue — Dashboard Urgencias SSMC

## 1. Configurar Supabase (base de datos)

1. Ve a [supabase.com](https://supabase.com) → **New project**
2. Elige nombre, contraseña y región (ej: South America)
3. En el panel, ve a **SQL Editor → New query**
4. Pega y ejecuta todo el contenido de `supabase_schema.sql`
5. Ve a **Project Settings → API** y copia:
   - `Project URL` → es tu `VITE_SUPABASE_URL`
   - `anon public` key → es tu `VITE_SUPABASE_ANON_KEY`

---

## 2. Subir el código a GitHub

```bash
git init
git add .
git commit -m "feat: migración Firebase → Supabase + deploy Netlify"
git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
git push -u origin main
```

---

## 3. Desplegar en Netlify

1. Ve a [netlify.com](https://netlify.com) → **Add new site → Import from Git**
2. Conecta tu repositorio de GitHub
3. Configuración de build (ya viene en `netlify.toml`):
   - Build command: `npm run build`
   - Publish directory: `dist`
4. Antes de hacer deploy, ve a **Site settings → Environment variables** y agrega:

   | Variable | Valor |
   |---|---|
   | `VITE_SUPABASE_URL` | `https://xxxx.supabase.co` |
   | `VITE_SUPABASE_ANON_KEY` | `eyJ...` |

5. Haz clic en **Deploy site** ✅

---

## ✅ Verificación

- La app debe cargar sin errores en consola
- Los datos se guardan en Supabase → Table Editor
- Los cambios aparecen en tiempo real en múltiples ventanas

