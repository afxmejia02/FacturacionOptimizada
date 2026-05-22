# Exportar `mapa-de-cargos` seguro a GitHub (Instrucciones en español)

Este documento explica cómo exportar la carpeta `mapa-de-cargos` a un repositorio remoto sin subir datos sensibles (por ejemplo la carpeta `docs/`).

Pre-requisitos (en tu máquina):
- Git instalado (https://git-scm.com/downloads)
- Acceso a línea de comandos (PowerShell en Windows)
- Opcional: Token de acceso (PAT) para repositorios privados en GitHub

Archivos añadidos por esta guía:
- `.gitignore` — excluye `docs/` y archivos sensibles.
- `push_to_github.ps1` — script PowerShell para inicializar y subir el repo.

Pasos rápidos:

1) Abre PowerShell y navega a la carpeta:

```powershell
cd C:\Users\andres.mejia\venv1.2\mapa-de-cargos
```

2) Revisa el `.gitignore` para confirmar qué se excluirá:

```powershell
cat .gitignore
```

3) Inicializa y sube usando el script (proporciona la URL remota):

```powershell
.\push_to_github.ps1 -RemoteUrl 'https://github.com/afxmejia02/VerificacionNominaTransferencia.git'
```

Notas de seguridad:
- La carpeta `docs/` está en `.gitignore` y no se subirá.
- Si hay archivos sensibles fuera de `docs/`, muévelos o añádelos a `.gitignore` antes de `git add`.
- Para repos privados, usa un PAT en lugar de tu contraseña.

Si prefieres usar SSH, reemplaza la URL HTTPS por la URL SSH y configura claves SSH en GitHub.

Problemas frecuentes:
- "git: command not found": instala Git en tu sistema.
- Errores de autenticación: crea un PAT (Settings → Developer settings → Personal access tokens) y utiliza un helper de credenciales o el prompt del OS.

Contacto:
Si necesitas que realice el push por ti, autoriza el acceso al host o comparte el PAT temporalmente (no recomendado por seguridad). En su lugar puedo guiarte paso a paso.
