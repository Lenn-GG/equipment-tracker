# G&G Equipment Requests

A phone-friendly web app for G&G Door Products crews to:

1. **Request** equipment for a job (company tools from a pick list, rentals, materials)
2. Get it **approved or denied** by a superintendent
3. Mark it **received on site** (condition, quantity, notes)
4. **Release** it from the site
5. Have the **warehouse manager check it back in**

Crews can also log **blocks, pallets and crates** from preinstall going back to the warehouse.

People sign in with their **Microsoft 365** accounts. Requests are stored in two **SharePoint lists**, so the data stays in G&G's Microsoft 365 tenant. The app itself is static files hosted free on **GitHub Pages**.

---

## One-time setup (about 15 minutes)

### Step 1: Register the app in Microsoft Entra ID

You need a Microsoft 365 admin account for this step.

1. Go to <https://entra.microsoft.com>, then **Applications → App registrations → New registration**.
2. **Name:** `G&G Equipment Requests`
3. **Supported account types:** *Accounts in this organizational directory only*.
4. **Redirect URI:** choose **Single-page application (SPA)** and enter your GitHub Pages address, for example
   `https://YOUR-GITHUB-USERNAME.github.io/gg-equipment/`
   (keep the trailing slash).
5. Click **Register**.
6. On the app's **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID**.
7. Open **API permissions → Add a permission → Microsoft Graph → Delegated permissions** and add:
   - `User.Read` (usually already there)
   - `Sites.ReadWrite.All`
   - `Sites.Manage.All` (used once, to create the SharePoint lists)
8. Click **Grant admin consent for G&G Door Products** and confirm.

### Step 2: Fill in `config.js`

Edit `config.js` in this repo:

| Setting | What to put |
|---|---|
| `clientId` | Application (client) ID from step 1 |
| `tenantId` | Directory (tenant) ID from step 1 |
| `sharePointSite` | The SharePoint site to keep the lists on, e.g. `https://ggdoor.sharepoint.com/sites/Operations` |
| `admins` | Emails that are always superintendents (you) |

Commit the change. GitHub Pages redeploys within a minute or two. None of these values are secret.

### Step 3: Give people access to the SharePoint site

Everyone who uses the app needs **Member (Edit)** access to the SharePoint site you picked. The simplest way is a site like "Operations" that your crew already belongs to.

### Step 4: First run

1. Open the app and **Sign in with Microsoft**.
2. The app offers to **Create SharePoint lists**. Click it. This makes `EquipmentRequests` and `EquipmentConfig` on the site and loads the default tool list.
3. Open **Settings** to add superintendents and warehouse managers by email, and to edit the tool list.

### Step 5: Put it on crew phones

Send the link. On iPhone use **Share → Add to Home Screen**; on Android use **⋮ → Add to Home screen**.

---

## Roles

| Role | Who | Can do |
|---|---|---|
| Superintendent | `admins` in `config.js`, plus emails listed in Settings | Approve / deny, edit settings, delete requests, check items in (backup) |
| Warehouse manager | Emails listed in Settings | Check items in at the warehouse, edit pick lists |
| Crew | Anyone else with access to the SharePoint site | Request, receive, release, send back blocks and pallets |

Roles are enforced by the app's screens. SharePoint permissions decide who can use the app at all.

## How data is stored

Each record is one item in the `EquipmentRequests` list:

- `Title`: request number (`EQ-YYMMDD-XXX` for requests, `RT-YYMMDD-XXX` for returns)
- `Status`: `pending`, `approved`, `denied`, `received`, `released` (headed to warehouse), `returned` (back in warehouse) or `consumed` (materials used up)
- `Payload`: JSON with the full history (who, when in Pacific time, condition, notes)

You can filter by `Status` in SharePoint or export the list to Excel for reports. Settings live in `EquipmentConfig` (`catalog` and `roles` items).

Open screens check for changes every 20 seconds and whenever the app comes back to the foreground.

## Files

- `index.html`: layout and styles
- `app.js`: app logic (Microsoft sign-in, SharePoint reads and writes)
- `config.js`: your tenant settings
- `vendor/msal-browser.min.js`: Microsoft's sign-in library (MIT license)
