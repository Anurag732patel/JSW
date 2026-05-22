# JSW Township DPR Dashboard

Welcome! This is the repository for the JSW Township Daily Progress Report (DPR) Dashboard. It's a full-stack tool built to help the JSW team parse Excel spreadsheets and instantly visualize township metrics like guest house occupancy, canteen meals, fleet movement, and maintenance tasks.

The frontend is built with **React** and **Vite**, and the backend is built with **Node.js, Express, and MongoDB**. They both run together smoothly on a single host.

---

## 🛠️ The Tech Stack & Style

- **Frontend**: React (Vite) + Recharts for clean, interactive charts.
- **Backend**: Node.js + Express.js.
- **Database**: MongoDB (configured for MongoDB Atlas in production).
- **Styling**: Custom Material Design-inspired dark theme with responsive CSS.
- **Deployments**: Configured to run on **Render** (completely free tier friendly).

---

## 🚀 Key Features

- **Upload Excel Spreadsheets**: Drop in a township DPR workbook (`.xlsx`, `.csv`), and the backend automatically parses and updates all indicators.
- **Interactive Visualizations**: View dynamic charts showing canteen meals, guest house occupancy rates, maintenance requests, and vehicle mileage.
- **Historical Reports**: Select any past date from the calendar dropdown to fetch and inspect historical data.
- **Automatic Database Seeding**: On the very first startup, the server automatically checks if the database is empty and creates the default admin user.

---

## 📁 Project Directory Layout

Here is how the codebase is organized:

```text
JSW/
├── backend/                  # Express server & database logic
│   ├── models/               # Mongoose schemas (User & DprReport)
│   ├── routes/               # API endpoints (Auth & DPR reports)
│   ├── utils/                # The excel parsing script
│   └── server.js             # Backend entry point (serves both API & built frontend)
├── src/                      # React frontend
│   ├── components/           # Protected routes & UI layout panels
│   ├── context/              # Authentication state provider
│   ├── pages/                # Pages (Dashboard, Main Dashboard, Login)
│   └── index.css             # Main stylesheet & theme colors
├── render.yaml               # 1-Click Render blueprint configuration
├── vite.config.js            # Frontend build configs & API proxy settings
└── package.json              # Project scripts & global settings
```

---

## 💻 Running the Project Locally

### 1. Set Up the Backend Environment
Go into the `backend` folder and create a `.env` file:
```env
PORT=5000
MONGODB_URI=mongodb://127.0.0.1:27017/jsw_db
JWT_SECRET=choose_any_long_random_string_here
```

### 2. Install & Start Everything

#### Start the Server:
```bash
cd backend
npm install
npm start
```

#### Start the Client (Vite Dev Server):
In a separate terminal in the root directory:
```bash
npm install
npm run dev
```
Open **[http://localhost:5173](http://localhost:5173)** in your browser. Any requests to `/api` will automatically proxy to the backend running on port `5000`.

---

## 🌐 Deploying to Render (Free Tier)

This app is configured to run on Render's free tier with a single click. Both the frontend and backend run on the same service to keep costs at zero.

### 1. Update MongoDB Atlas Network Access
Render's free tier uses dynamic IP addresses, so you need to whitelist all IPs in MongoDB Atlas:
- Go to your MongoDB Atlas dashboard -> **Network Access**.
- Click **Add IP Address** -> click **Allow Access from Anywhere** (`0.0.0.0/0`) -> **Confirm**.

### 2. Launch on Render
1. Log into your **[Render Dashboard](https://dashboard.render.com/)**.
2. Click **New** -> **Blueprint**.
3. Select your **`JSW`** repository.
4. Paste your **`MONGODB_URI`** connection string when prompted in the environment variables form.
5. Click **Approve / Deploy**.

*Render will automatically use Node 22, install all dependencies, build the React app, and launch the server. It will also auto-seed the default admin credentials on the first run!*

---

## 🔑 Default Administrator Login

Once the application is up and running, you can log in using these default credentials:
- **Username**: `admin`
- **Password**: `jsw@2024`

*(If you ever need to reset the database or seed it again, you can run `npm run seed` in your local `backend` folder.)*
