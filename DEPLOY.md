# APSG Report — Fly.io Deployment Guide

## What's Changed

### 1. New Files Added
| File | Purpose |
|------|---------|
| `Dockerfile` | Container build for Fly.io |
| `fly.toml` | Fly.io app configuration |
| `.dockerignore` | Exclude unnecessary files from Docker build |

### 2. app.py Changes
- **Dynamic PORT** from `os.environ.get("PORT", 8080)` — already present, now binds `0.0.0.0`
- **Persistent volume support**: DB, uploads, and outputs now use `/app/data/` when the Fly.io volume is mounted
- **Full UI overhaul**: All 7 pages (Login, Dashboard, PPT Rejection, Excel Rejection, Photo Merge, Daily Report, Admin) upgraded to modern SaaS design

---

## Fly.io Deployment Steps

### Prerequisites
```bash
# Install flyctl
curl -L https://fly.io/install.sh | sh

# Login
flyctl auth login
```

### Deploy
```bash
# From the fpa_report/ directory:
fly launch          # First time — creates app, prompts for name/region
fly deploy          # Subsequent deploys
```

When prompted during `fly launch`:
- **App name**: `apsg-report` (or your preferred name)
- **Region**: `sin` (Singapore) or nearest to you
- **PostgreSQL**: No
- **Redis**: No

### Create Persistent Volume (recommended)
```bash
fly volumes create apsg_data --size 1 --region sin
```
This stores the SQLite DB, uploads, and outputs across restarts.

### Set Secrets
```bash
fly secrets set SECRET_KEY="$(openssl rand -hex 32)"
```

### Domain Setup (apsgreport.site)
```bash
# Add custom domain
flyctl certs add apsgreport.site
flyctl certs add www.apsgreport.site

# Get Fly.io IP for DNS
flyctl ips list
```

**Namecheap DNS Records:**
| Type | Host | Value |
|------|------|-------|
| A | @ | `<Fly.io IPv4 from flyctl ips list>` |
| AAAA | @ | `<Fly.io IPv6 from flyctl ips list>` |
| CNAME | www | `apsg-report.fly.dev` |

Wait 10–60 min for DNS propagation, then verify:
```bash
flyctl certs check apsgreport.site
```

---

## UI Changes Summary

### Design System
- **Colors**: Dark navy base (`#080C1A`) + Indigo/Purple/Cyan gradient accents
- **Font**: Inter (body) + Poppins (headings/brand)
- **Cards**: Glass-morphism with `backdrop-filter: blur(16px)`, 16–18px border-radius
- **Background**: Layered radial gradients + subtle grid lines

### Per-Page Changes
| Page | Key Changes |
|------|------------|
| **Login** | New logo tile, gradient background mesh, shimmer button, cleaner form |
| **Dashboard** | Gradient hero title, card hover with glow + lift, arrow indicators |
| **All sub-pages** | Modern top bar (56px), improved upload zones, gradient buttons |
| **Upload zones** | Drag-over highlight, file name display, green border on success |
| **Buttons** | Gradient primary, shimmer hover effect, micro-interactions |

---

## Performance Notes (Free Tier)
- `auto_stop_machines = true` — machine sleeps when idle (saves credits)
- `auto_start_machines = true` — wakes on first request (~2s cold start)
- Workers set to 2 (gunicorn) — sufficient for internal tool usage
- No heavy animations — only hover micro-interactions (CSS only, no JS animation loops)
- `backdrop-filter` used only on cards, not full-page (GPU-friendly)

---

## Troubleshooting

```bash
# View logs
fly logs

# SSH into machine
fly ssh console

# Check app status
fly status

# Restart
fly machine restart
```
