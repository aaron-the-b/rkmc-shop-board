# RKMC Shop Board

Interactive floor plans for RKMC1 and RKMC2: production lines, stations, skids, serial numbers and job access, saved to the RKMC Shop Board SharePoint site.

Live page: https://aaron-the-b.github.io/rkmc-shop-board/

## How it fits together

| File | What it does |
| --- | --- |
| `index.html` | Page layout and styles |
| `app.js` | The board itself: map, panels, editing |
| `m365.js` | Microsoft sign-in, loading from SharePoint, saving changes back |
| `config.js` | Which Entra app and SharePoint site to use (no secrets) |

**No RK data lives in this repository.** Shop drawings, layouts, roles and job teams are in `ShopBoard/board-config.json` in the site's Documents library. Jobs, skids, serials, people, groups and access are in the SharePoint lists of the same names. The page shows nothing until someone signs in with an RK account.

## SharePoint setup

Lists (from `ShopBoard-List-Templates.xlsx`): Jobs, Skids, SerialRequirements, SerialRecords, ActivityLog, Groups, JobAccess, UsersRoles. Document library: SkidPhotos.

Add one column by hand: **Skids → BoardData** (Multiple lines of text, plain text). It holds each skid's QC checklist answers and history. Without it, positions and serials still save but QC answers don't.

The first person to sign in while UsersRoles is empty becomes the administrator.

## Publishing

Every push to `main` publishes the page through GitHub Pages (Settings → Pages → Source: GitHub Actions).
