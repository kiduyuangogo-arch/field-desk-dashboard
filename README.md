# Field Desk — survey analytics dashboard

A browser-based dashboard for exploring CSV/Excel exports from Survey123,
KoboToolbox, Google Forms, or any similar tool. Upload a file, and it
auto-detects field types, builds summary charts, and gives you a searchable
data table. Nothing is uploaded anywhere — all parsing happens in the
browser.

## Run it locally

Requires [Node.js](https://nodejs.org) 18 or later.

```bash
npm install
npm run dev
```

Then open the URL it prints (usually http://localhost:5173).

## Deploy to Netlify

**Option A — drag and drop (quickest)**
```bash
npm install
npm run build
```
This creates a `dist` folder. Go to [app.netlify.com/drop](https://app.netlify.com/drop)
and drag the `dist` folder in. Done — you'll get a live URL immediately.

**Option B — connect a Git repo (recommended for ongoing use)**
1. Push this folder to a GitHub/GitLab/Bitbucket repo.
2. In Netlify: **Add new site → Import an existing project**, pick the repo.
3. Netlify will auto-detect the settings from `netlify.toml`:
   - Build command: `npm run build`
   - Publish directory: `dist`
4. Click **Deploy**. Every future push updates the live site automatically.

## Project structure

```
field-desk/
├── index.html          Vite entry HTML
├── src/
│   ├── main.jsx         Mounts the React app
│   ├── App.jsx           The dashboard itself
│   └── index.css        Tailwind setup
├── package.json         Dependencies and scripts
├── vite.config.js       Build tool config
├── tailwind.config.js   Styling config
└── netlify.toml         Tells Netlify how to build/publish
```

## Notes

- Data resets on page reload — nothing is persisted between sessions yet.
- To add live API connections (Survey123 REST API, KoboToolbox API) later,
  you'd add a small serverless function (Netlify Functions work well here)
  to handle authentication, since browsers can't call those APIs directly.
