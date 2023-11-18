# var_proj

[![var_proj](/assets/media/logo.png)](https://var_proj/)


## STEPS


### Repositories


#### Backup

- [Github new acount](https://github.com/join)
  - User name: `var_user`
  - Email: `var_user@seacomoseo.com`
  - Password: `var_pass`
  - Verify and exit
- [New repository](https://github.com/new)
  - `Repository name` = `var_proj`
  - `Private` = check
  - `Create repository`
- [Generate new token](https://github.com/settings/tokens/new)
  - `Note` = `gitlab-mirror`
  - `Expiration` = `No expiration`
  - `Select scopes` = `repo`
  - `Generate token`
  - Copy token (`ghp_**********`) for `Mirror repository` in `Password`


#### Main

- [Gitlab new acount](https://gitlab.com/users/sign_up)
  - User name: `var_user`
  - Email: `var_user@seacomoseo.com`
  - Password: `var_pass`
  - Note in `🔥 SanSoul/var_proj.md`
  - Verify
  - Role: `Other`
  - Check `Just me` and `Join a project`
  - `Continue`
- New or fork project
  - If [New project](https://gitlab.com/projects/new#blank_project)
  - If [Fork this project](https://gitlab.com/var_user/var_proj/-/forks/new)
  - `Project name` and `Project slug` = `var_proj`
  - `Visibility level` = `Private`
  - `Initialize repository with a README` = uncheck
- [Mirror repository](https://gitlab.com/var_user/var_proj/-/settings/repository#js-push-remote-settings)
  - `Add new`
  - `Git repository URL` = `https://var_user@github.com/var_user/var_proj.git`
  - `Password` = prev token from Github (`ghp_**********`)
  - `Mirror repository`
- [Members](https://gitlab.com/var_user/var_proj/-/project_members) ⏩ Add your email with role `Maintainer`
- Icons
  - [General Settings](https://gitlab.com/var_user/var_proj/edit) ⏩ `Project avatar`
  - [Profile](https://gitlab.com/-/profile) ⏩ `Public avatar`
- [Preferences](https://gitlab.com/-/profile/preferences)
  - `Language` `Spanish`
  - `First day of the week` `Moonday`
- [Repository](https://gitlab.com/var_user/var_proj) ⏩ `Clone ⏩ Visual Studio Code (HTTPS)` ⏩ select sansoul folder (not new folder)


### Hosting


#### If Netlify

- `netlify.toml + package.json` must be existed
- [Netlify new acount](https://app.netlify.com/signup/email)
  - User name: `var_user`
  - Email: `var_user@seacomoseo.com`
  - Password: `var_pass`
  - Verify and exit
- `Skip personalization ⏩ New site from Git ⏩ Gitlab ⏩ Authorize ⏩ Select repo ⏩ Deploy site ⏩ Site settings`
  - `Change site name ⏩ var_user` and copy in `data/config.yml ⏩ netlify_name`
  - Copy `API ID` in `data/config.yml ⏩ netlify_id`
  - [Build hooks](https://app.netlify.com/sites/var_user/settings/deploys#build-hooks) with = `build` name ⏩ copy and paste in `var_proj/var_proj.md` and [spreadsheet projects](https://docs.google.com/spreadsheets/d/1QIeT5VfUuXC3momIScZPf8ZPahV23tXVNshcf6DRipA/).
- [Profile Avatar](https://app.netlify.com/user/settings#profile) ⏩ `Edit settings` ⏩ `Avatar`
- [Connected Accounts](https://app.netlify.com/user/settings#connected-accounts) ⏩ `Edit settings` ⏩ `GitLab`: `Connect`


#### If Cloudflare Pages

- [Cloudflare new acount](https://dash.cloudflare.com/sign-up)
  - User name: `var_user`
  - Email: `var_user@seacomoseo.com`
  - Verify and exit
- Copy id by URL in `data/config.yml ⏩ cloudflare_id`
- [Cloudflare Pages new project with Gitlab](https://dash.cloudflare.com/?to=/:account/pages/new/provider/gitlab)
  - `Connect Gitlab` ⏩ `Authorize` ⏩ `Select a repository` ⏩ `Begin setup`
  - `Project name` = `var_user` and copy in `data/config.yml ⏩ cloudflare_project_name`
  - `Framework preset` = `Hugo`
  - `Build command` = `sh do build`
  - `Build output directory` = `public`
  - `Environment variables (advanced)` ⏩ `Add variable` ⏩ `HUGO_VERSION` = `0.92.2`
  - `Save and Deploy`
- [Language Preference](https://dash.cloudflare.com/profile#language-preference-card)


### CMS

- [`Gitlab ⏩ Edit profile ⏩ Access Token`](https://gitlab.com/-/profile/personal_access_tokens)
  - `Token name`: `netlify-identity`
  - `Expiration date`: remove
  - `Select scopes`: check `api`, `read_api`, `read_repository` and `write_repository`
  - `Create personal access token` ⏩ copy `our new personal access token`
- If Identity Access
  - [`Netlify ⏩ Site ⏩ Settings ⏩ Identity`](https://app.netlify.com/sites/var_user/settings/identity) ⏩ `Enable Identity`
    - [`Git Gateway`](https://app.netlify.com/sites/var_user/settings/identity#git-gateway) ⏩ `Edit settings` ⏩ Paste the gitlab code copied in `GitLab personal access token ⏩ Save`
    - [Registration preferences](https://app.netlify.com/sites/var_user/settings/identity#registration-preferences) ⏩ Edit settings ⏩ Invite only ⏩ Save
    - [`External providers`](https://app.netlify.com/sites/var_user/settings/identity#external-providers) ⏩ Add provider ⏩ Gitlab ⏩ Use default configuration ⏩ Enable GitLab
  - [`Netlify ⏩ Site ⏩ Identity`](https://app.netlify.com/sites/var_user/identity) ⏩ Invite users ⏩ email `var_user@seacomoseo.com` (and other if you want) ⏩ Check email from `no-reply@netlify.com` ⏩ Accept the invite ⏩ Add `/admin/` in the URL between `/` and `#` ⏩ Add password `var_pass`
- If Cloudflare Pages Access
  - ..........................................................


### Local

- If new project (not fork):
  - Copy [var_proj project files](https://github.com/var_user/var_proj)
  - `git submodule add https://github.com/lorensansol/sansoul.git themes/sansoul`
- Else if fork:
  - Download submódule theme files: `git submodule update --init --recursive`
- If you want use svg emojies:
  - `git submodule add https://github.com/lorensansol/sansoul-emojis.git layouts/partials/svg/emojis`
- Now you can see the run project in the browser with the `do server` comand
- `README.md` ⏩ edit baseURL ("var_proj") + delete steps bit by bit
- Design
  - You can edit any file with the same structure of theme sansoul
  - GENERAL
    - `config.yml`
    - `data/*.{yml,md}`
  - IMG + LOGO + FAVICON
    - `assets/media/` folder ⏩ [Compress image tool](https://compressor.io/)
      - `fondo.jpg`
      - `logo.svg`
      - `logo.png`
      - `icon.png`
      - `favicon.ico` ⏩ [Favicon converter tool](https://favicon.io/favicon-converter/)
    - Gitlab, Github, Netlify and Cloudflare Pages update icon (project and account)
  - FONTS
    - Fonts that not be in Google Fonts:
      - .otf/.ttf files in `assets/fonts` + rename
      - You need the `sansoul_tools` project in `../_tools` folder
      - Run `do fonts` comand
      - `Family` + `Style` + `Weight` must match in `config.yml ⏩ params.fonts` + `_fonts.scss` (and disable `params.fonts.google` if appropriate)
  - CONTENT
    - `content/*`
      - `blog/divisores.md` ⏩ remove
      - `admin` ⏩ `draft: true` and/or change params and content
      - SCRAP
        - Copy [this Spreadsheet Layout](https://docs.google.com/spreadsheets/d/1bJQaAFoBAwHhWz_WFRIHlkHoDmZZyT8H8NtdCnSTdaU)
        - First scrap with Screaming Frog and paste `url` and `title`
        - Second scrap with `IMPORTXML` formula in `content-start` tab
        - If need HTML content
          - Three scrap with `Web Scraper` chrome extension and paste in `content-scrap` tab
          - Copy `content-start` tab into `content-next` and get `body_code` (by `content-scrap`) with `VLOOKUP` formula
          - Copy `body_code` column to `body`, replace double quotes, [convert to Markdown](https://smalldev.tools/html-to-markdown-converter-online) and clean it
          - Paste in `file.md` in `./Downloads` and run `_tools/md-replaces.command`
          - Check `file_new.md` content:
            - headers: `^#`
            - images: `!\[.+?\)`
            - links: `(^|[^!])\[.+?\)`
            - internal links: `https?://(www\.)?example\.com`
          - Paste in sheet
        - Export to CSV like `./Downloads/file.csv` and run `_tools/csv-to-md.command`
        - Move files from `markdown_files` to `blog` project
  - HTML
    - `layouts/*`
    - `data/config.yml ⏩ custom_code.html_head.code`
    - `data/config.yml ⏩ custom_code.html_body.code`
  - CSS
    - `assets/css/*`
    - `assets/css/` ⏩ `_variables-custom.scss` + `_custom.scss`
    - `data/config.yml ⏩ custom_code.css.code`
  - JS
    - `assets/js/*`
    - `layouts/partials/sections/common/scripts.html`
    - `data/config.yml ⏩ custom_code.js.code`
  - IMG
    - `assets/media/*`
  - REDIRECTS
    - `assets/redirects.md`
  - ROBOTS
    - `assets/robots.md`
  - If Multilanguaje and Multihosting, add `cp ./public/[es|en]/404.html ./public/` in `netlify.toml ⏩ build.command`
  - Try in Safari and Firefox
  - Check in [W3 Validator](https://validator.w3.org/)
  - Check in [PageSpeed Insights](https://pagespeed.web.dev/)


### After client validate web


#### Domain

- If Netlify
  - [`Domain Management / settings`](https://app.netlify.com/sites/var_user/settings/domain)
  - `Add custom domain`
  - `Check DNS configuration` Copy
  - Add `DNS Records` copied from Netlify to Domain gestor:
    - From: `var_proj`
      DNS Record: `ALIAS`, `ANAME` or `flattened CNAME`
      To: `apex-loadbalancer.netlify.com`
    - From: `var_proj`
      DNS Record: `A`
      To: `75.2.60.5`
    - From: `www`
      DNS Record: `CNAME`
      To: `var_user.netlify.app.`
    - Maybe you need to eliminate the previous records with similar names
  - `Verify DNS configuration`
  - If it does not work after a while, try `Set as main domain` in the `www` version and also in te `nowww` version
- If Cloudflare Pages
  - [Custom domains](https://dash.cloudflare.com/?to=/:account/pages/view/var_user/domains)
  - `Set up a custom domains`
  - `var_proj`
  - `Continue`
  - `Activate domain` (if `Begin DNS transfer` end)
  - Repeat with `www.var_proj`
  - ...........................................................


#### Forms

- If Netlify Form
  - Don't need configure nothing! Build like you want in local or with CMS
  - [`Netlify ⏩ Site ⏩ Forms ⏩ Form Notifications`](https://app.netlify.com/sites/var_user/settings/forms#form-notifications) ⏩ `Add notification ⏩ Email notification ⏩ Email to Notify`
    - `Email to notify` = Emails of collaborators that want receive submissions
    - `Custom email subject line` = `Formulario de contacto de var_proj`
    - `Save`
  - Submissions: [`Netlify site ⏩ Forms`](https://app.netlify.com/sites/var_user/forms)
- If Cloudflare Workers
  - ...........................................................
- [formsubmit.co](https://formsubmit.co/)
- If Google Form: [Tutorial](https://seacomoseo.com/instrucciones/#google-forms)


#### [Google Analytics](https://analytics.google.com/)

- `Admin ⏩ Libre acount ⏩ New property ⏩ ...` copy ID
- `data/config.yml ⏩ google_analytics` ⏩ paste ID
- `Ajustes de datos`
  - `Recogida de datos`
    - `Recogida de datos de Google signals ⏩ Empezar`
    - `Consentimiento de recogida de datos de usuario` ⏩ Check
  - `Conservación de datos ⏩ Conservación de datos de eventos ⏩ 14 meses ⏩ Guardar`
- `Conversiones ⏩ Nuevo evento de conversión ⏩ Nombre de evento nuevo` ⏩ add `contact_click` and `contact_form_submit`


#### [Google Search Console](https://search.google.com/search-console)

- Add property
- Verify versions ⏩ `data/config.yml`
  - `google_analytics`
  - `google_site_verification`
  - `google_file_verification`
  - DNS:
    From: `@`
    DNS Record: `TXT`
    To: `google-site-verification=[google_site_verification]`
- Link with Analytics
- Add sitemap.xml


#### [Google My Business](https://business.google.com/)

- `Add company ⏩ ...` ⏩ whait 13 days to receive postal and insert code to verify


#### [Disqus](https://disqus.com/)

- `data/config.yml ⏩ disqus`


#### Collaborators

- [Google Analytics](https://analytics.google.com/) ⏩ `Admin ⏩ Libre acount ⏩ Site ⏩ Property access management ⏩ Add users` ⏩ Add emails of collaborators with role `Reader` or `Admin`.
- [Google Search Console](https://search.google.com/search-console) ⏩ `Site ⏩ Settings ⏩ Users and permissions ⏩ Add user` Add emails of collaborators with `Full` permission
- [Google My Business](https://business.google.com/)
  - `Site ⏩ Users ⏩ Add users` ⏩ Add emails of collaborators with role `Owner`


##### Services Layout

1. [Servicios var_proj](https://drive.google.com/file/d/1trq28fMfEVwoZOk4ue0tJzAJDZtj64BK) ⏩ `File ⏩ Make a copy` ⏩ Select client directory.
1. Change the info.
1. `Share` ⏩ Add emails of collaborators with `Editor` permission.


##### Delivery

Send to all collaborators next:

```
*ENTREGA WEB var_proj:* https://seacomoseo.com/entrega/
```

## GADS

### GA4

- Vincular:
  - `Herramientas y configuración`
  - `Cuentas vinculadas`
  - `Google Analytics (GA4) & Firebase ⏩ Detalles`
  - Buscar proyecto y ⏩ `Vincular`
- Importar conversiones:
  - `Herramientas y configuración`
  - `Medición`
  - `Conversiones`
  - `Nueva acción de conversión`
  - `Importar`
  - `Propiedades de Google Analytics 4`
  - `Web`
  - `Continuar`
  - Seleccionar `contact_click` y `contact_form_submit`
  - `Importar y continuar`
- Cambiar a `Maximizar conversiones`
  - `Campañas`
  - `⚙️`
  - `Puja`
  - `Cambiar estrategia de puja`
  - `¿En qué quieres centrarte? ⏩ Conversiones`
