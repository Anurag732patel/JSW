import { useState, useRef, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { motion, AnimatePresence } from 'framer-motion';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell,
  LineChart, Line, AreaChart, Area,
} from 'recharts';

// ─── DEFAULT DATA (from DPR Sheets 23/09/2025) ──────────────────────────────

// Sheet 1 — F&B: Manpower
const defaultManpowerFB = [
  { name: 'On Roll', actual: 11, present: 11 },
  { name: 'Associates', actual: 62, present: 53 },
  { name: 'Rashmi (JBN)', actual: 17, present: 16 },
  { name: 'Rashmi (JBC)', actual: 161, present: 135 },
  { name: 'Infinzy', actual: 7, present: 4 },
  { name: 'Sparkal', actual: 27, present: 25 },
];

// Sheet 1 — Guest House Occupancy
const defaultGuestHouseData = [
  { name: 'IB River', total: 76, occupied: 19, vacant: 57 },
  { name: 'VIP', total: 33, occupied: 9, vacant: 24 },
  { name: 'New GET', total: 42, occupied: 22, vacant: 20 },
];

// Sheet 1 — Meal Count (top canteens)
const defaultMealData = [
  { name: 'NG V.I.P', breakfast: 197, lunch: 484, dinner: 148 },
  { name: 'NG VIP Mail', breakfast: 107, lunch: 166, dinner: 114 },
  { name: 'WRM', breakfast: 97, lunch: 163, dinner: 71 },
  { name: 'CRM', breakfast: 67, lunch: 114, dinner: 27 },
  { name: 'CSP', breakfast: 98, lunch: 122, dinner: 48 },
  { name: 'GET Mail', breakfast: 80, lunch: 200, dinner: 106 },
  { name: 'V.I.P', breakfast: 31, lunch: 41, dinner: 19 },
  { name: 'PELLET', breakfast: 38, lunch: 82, dinner: 27 },
];

// Sheet 1 — Vehicle Km
const defaultVehicleData = [
  { name: 'MG-EV 9304', km: 115 },
  { name: 'MG-EV 9346', km: 138 },
  { name: 'MG-EV 9357', km: 97 },
  { name: 'MG-EV 9203', km: 76 },
  { name: 'MG-EV 9147', km: 76 },
  { name: 'MG-EV 9398', km: 83 },
  { name: 'MG-EV 1797', km: 83 },
  { name: 'MG-EV 9273', km: 21 },
  { name: 'Bolero', km: 41 },
];

// Sheet 2 — Maintenance Manpower
const defaultManpowerMaint = [
  { name: 'On Roll', actual: 14, present: 11 },
  { name: 'Associates', actual: 51, present: 41 },
  { name: 'Kanha Ent.', actual: 11, present: 11 },
  { name: 'Maa Sarala', actual: 17, present: 15 },
  { name: 'Samaleswari', actual: 15, present: 13 },
  { name: 'Global Elevator', actual: 2, present: 2 },
];

// Sheet 2 — Complaints by category
const defaultComplaintsPie = [
  { name: 'Civil', value: 13, color: '#c0293a' },
  { name: 'Seepage', value: 26, color: '#e07a5f' },
  { name: 'Carpentry', value: 12, color: '#2a9d8f' },
  { name: 'Plumbing', value: 2, color: '#0b4d8c' },
  { name: 'Electrical', value: 3, color: '#475569' },
];

// Sheet 2 — Complaints status
const defaultComplaintsStatus = [
  { name: 'Civil', completed: 0, inProgress: 13, pending: 0 },
  { name: 'Seepage', completed: 0, inProgress: 26, pending: 0 },
  { name: 'Carpentry', completed: 12, inProgress: 0, pending: 0 },
  { name: 'Plumbing', completed: 2, inProgress: 0, pending: 0 },
  { name: 'Electrical', completed: 3, inProgress: 0, pending: 0 },
];

// Sheet 2 — JBC Paint work progress (%)
const defaultPaintProgress = [
  { name: 'Pappi Devi\nInt. Wall', target: 5000, done: 4766.73 },
  { name: 'Pappi Devi\nExt. Wall', target: 500, done: 135.84 },
  { name: 'Manoj Ent.\nInt. Wall', target: 5000, done: 5089.73 },
  { name: 'Manoj Ent.\nExt. Wall', target: 500, done: 405.85 },
];

// Sheet 2 — Complaints List (Default / Sample data)
const defaultComplaintsList = [
  {
    department: 'Civil',
    jobNo: '5064',
    description: 'Entrance grill painting work start at M3 building ground floor',
    completedDetails: '',
    pendingDetails: 'work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '5317',
    description: 'Vehicle parking painting near L & M block',
    completedDetails: '',
    pendingDetails: 'Work stopped for certain time due to rain',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'N & O Block parking painting',
    completedDetails: '',
    pendingDetails: 'Work stopped for certain time due to rain',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'E3-701 Putty & painting work',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'NEW GET D-914 Putty & painting work',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'A & C Block no parking painting',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'D-001 Putty & painting',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'E3-603 Kitchen sink repairing',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'C-203(Vacant room) Both bathroom wall tiles fixing',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'NEW GET CF-102(Vacant room) Full flat painting',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'H-604 Putty & painting',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'M3-604(Vacant room) Full flat painting',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Civil',
    jobNo: '',
    description: 'Temple lord Vishwakarma stage marble work',
    completedDetails: '',
    pendingDetails: 'Work in progress',
    feedback: ''
  },
  {
    department: 'Plumbing',
    jobNo: '',
    description: 'FU1-005 Kitchen sink waste coupling damage',
    completedDetails: 'New installed',
    pendingDetails: '',
    feedback: 'Outstanding'
  },
  {
    department: 'Plumbing',
    jobNo: '',
    description: 'E3-601 Kitchen sink cock damage',
    completedDetails: 'New installed',
    pendingDetails: '',
    feedback: 'Outstanding'
  },
  {
    department: 'Seepage',
    jobNo: '5888',
    description: 'E2-701 Balcony ceiling seepage',
    completedDetails: '',
    pendingDetails: 'Work will start as soon as possible',
    feedback: ''
  },
  {
    department: 'Seepage',
    jobNo: '6769',
    description: 'E2-804 Hall ceiling seepage',
    completedDetails: '',
    pendingDetails: 'Work will start as soon as possible',
    feedback: ''
  },
  {
    department: 'Seepage',
    jobNo: '',
    description: 'J-302 Master bedroom ceiling seepage',
    completedDetails: '',
    pendingDetails: 'Work will start as soon as possible',
    feedback: ''
  },
  {
    department: 'Seepage',
    jobNo: '8893',
    description: 'IB Basement floor seepage',
    completedDetails: '',
    pendingDetails: 'Work under progress',
    feedback: ''
  }
];

// ─── Section Definitions ────────────────────────────────────────────────────
const SECTIONS = [
  { id: 'guest-house', label: 'Guest House Occupancy', icon: '🏨' },
  { id: 'meal-report', label: 'Canteen Meal Report', icon: '🍽️' },
  { id: 'fb-manpower', label: 'F&B Dept. Manpower', icon: '👷' },
  { id: 'vehicle-fleet', label: 'Fleet Vehicle Movement', icon: '🚗' },
  { id: 'complaints', label: 'Maintenance Complaints', icon: '🔧' },
  { id: 'paint-work', label: 'Paint Work Progress', icon: '🎨' },
  { id: 'maint-manpower', label: 'Maintenance Manpower', icon: '🏗️' },
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
const TOOLTIP_STYLE = {
  contentStyle: { background: '#1e293b', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 8, color: '#f1f5f9' },
  labelStyle: { color: '#94a3b8', fontSize: 12 },
  itemStyle: { color: '#f1f5f9' },
};

// ─── Main Component ───────────────────────────────────────────────────────────
export default function MainDashboard() {
  const { user, logout } = useAuth();
  const navigate = useNavigate();
  const fileInputRef = useRef(null);
  const profileMenuRef = useRef(null);
  const dateDropdownRef = useRef(null);

  const [activeSection, setActiveSection] = useState('guest-house');
  const [sectionsPanelOpen, setSectionsPanelOpen] = useState(false);
  const [importedData, setImportedData] = useState(null);
  const [importFileName, setImportFileName] = useState('');
  const [profileDropdownOpen, setProfileDropdownOpen] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analyzingStep, setAnalyzingStep] = useState('');

  // ─── Historical Navigation & Ingestion States ───
  const [availableDates, setAvailableDates] = useState([]);
  const [selectedDate, setSelectedDate] = useState('');
  const [dateDropdownOpen, setDateDropdownOpen] = useState(false);
  const [showDateModal, setShowDateModal] = useState(false);
  const [pendingFile, setPendingFile] = useState(null);
  const [inputDate, setInputDate] = useState('');

  const [showComplaintsTable, setShowComplaintsTable] = useState(false);

  useEffect(() => {
    setShowComplaintsTable(false);
  }, [activeSection]);

  // Close dropdowns when clicking outside
  useEffect(() => {
    function handleClickOutside(event) {
      if (profileMenuRef.current && !profileMenuRef.current.contains(event.target)) {
        setProfileDropdownOpen(false);
      }
      if (dateDropdownRef.current && !dateDropdownRef.current.contains(event.target)) {
        setDateDropdownOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Fetch report data for a specific date
  const fetchReportForDate = async (date) => {
    try {
      const res = await fetch(`/api/dpr/report/${date}`);

      let result;
      const contentType = res.headers.get("content-type");
      if (contentType && contentType.includes("application/json")) {
        result = await res.json();
      } else {
        const text = await res.text();
        const cleanText = text.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
        alert(`Server Error: ${cleanText || `Status ${res.status}`}`);
        return;
      }

      if (result.success) {
        setImportedData(result.data);
        setImportFileName(result.filename);
        setSelectedDate(date);
      } else {
        alert(`Failed to load report: ${result.message}`);
      }
    } catch (err) {
      console.error("Failed to fetch report for date:", err);
      alert("Error contacting JSW backend server. Verify connection.");
    }
  };

  // Fetch available dates on mount (without automatically loading a report data)
  useEffect(() => {
    const loadInitialData = async () => {
      try {
        const res = await fetch('/api/dpr/dates');

        let result;
        const contentType = res.headers.get("content-type");
        if (contentType && contentType.includes("application/json")) {
          result = await res.json();
        } else {
          console.error("Non-JSON response fetched for dates:", res.status);
          return;
        }

        if (result.success && result.dates) {
          setAvailableDates(result.dates);
          // Auto-load on mount is disabled to show a clean upload prompt initially.
          // Users can select any historical date from the calendar dropdown or upload a new spreadsheet.
        }
      } catch (err) {
        console.error("Failed to load initial dates:", err);
      }
    };
    loadInitialData();
  }, []);

  // ─── Compute KPI values ──────────────────────────────────────────
  const mealData = importedData?.mealData || [];
  const guestHouseData = importedData?.guestHouseData || [];
  const vehicleData = importedData?.vehicleData || [];
  const manpowerFB = importedData?.manpowerFB || [];
  const manpowerMaint = importedData?.manpowerMaint || [];
  const complaintsPie = importedData?.complaintsPie || [];
  const complaintsStatus = importedData?.complaintsStatus || [];
  const paintProgress = importedData?.paintProgress || [];
  const complaintsList = importedData?.complaintsList || defaultComplaintsList;

  const totalBreakfast = mealData.reduce((s, m) => s + (m.breakfast || 0), 0);
  const totalLunch = mealData.reduce((s, m) => s + (m.lunch || 0), 0);
  const totalDinner = mealData.reduce((s, m) => s + (m.dinner || 0), 0);
  const totalMeals = totalBreakfast + totalLunch + totalDinner;
  const totalKm = vehicleData.reduce((s, v) => s + (v.km || 0), 0);
  const totalRooms = guestHouseData.reduce((s, g) => s + (g.total || 0), 0);
  const occupiedRooms = guestHouseData.reduce((s, g) => s + (g.occupied || 0), 0);
  const totalComplaints = complaintsPie.reduce((a, b) => a + b.value, 0);

  const currentSection = SECTIONS.find(s => s.id === activeSection);

  // ─── Import Handlers (Database-backed Ingestion) ───────────────────
  const handleFileSelect = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setPendingFile(file);
    // Pre-fill date input with today's date in local YYYY-MM-DD
    const today = new Date();
    const offset = today.getTimezoneOffset();
    const localDate = new Date(today.getTime() - (offset * 60 * 1000));
    const formattedDate = localDate.toISOString().split('T')[0];
    setInputDate(formattedDate);
    setShowDateModal(true);
    e.target.value = ''; // Reset file input
  };

  const handleImportConfirm = async () => {
    if (!pendingFile || !inputDate) return;
    setShowDateModal(false);
    setImportFileName(pendingFile.name);
    setIsAnalyzing(true);
    setAnalyzingStep('Est. Secure Connection with JSW Backend...');

    try {
      const formData = new FormData();
      formData.append('file', pendingFile);
      formData.append('reportDate', inputDate);
      formData.append('uploadedBy', user?.name || 'System Administrator');

      const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

      await delay(400);
      setAnalyzingStep('Uploading spreadsheet (multipart/form-data)...');

      const responsePromise = fetch('/api/dpr/import', {
        method: 'POST',
        body: formData,
      });

      await delay(500);
      setAnalyzingStep('Scanning workbook sheets in Node.js memory buffer...');

      await delay(500);
      setAnalyzingStep('Running mapping algorithms on Excel columns...');

      await delay(400);
      setAnalyzingStep('Validating schema metrics & running Recharts parser...');

      const response = await responsePromise;

      let result;
      const contentType = response.headers.get("content-type");
      if (contentType && contentType.includes("application/json")) {
        result = await response.json();
      } else {
        const text = await response.text();
        // Remove HTML tags for clean alert display if it is an HTML error page
        const cleanText = text.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
        throw new Error(cleanText || `Server returned status ${response.status}`);
      }

      if (!response.ok) {
        throw new Error(result?.message || `Upload failed with status ${response.status}`);
      }

      if (result.success) {
        await delay(300);
        setAnalyzingStep('Finalizing database upsert...');
        await delay(200);
        setImportedData(result.data);
        setSelectedDate(inputDate);

        // Refresh available dates list and highlight new report
        const datesRes = await fetch('/api/dpr/dates');
        const datesResult = await datesRes.json();
        if (datesResult.success && datesResult.dates) {
          setAvailableDates(datesResult.dates);
        }
      } else {
        throw new Error(result.message || 'Excel processing failed');
      }
    } catch (err) {
      console.error('Import error:', err);
      alert(`Import Failed: ${err.message || 'Server error. Please ensure your backend is running.'}`);
    } finally {
      setIsAnalyzing(false);
      setAnalyzingStep('');
      setPendingFile(null);
    }
  };

  // ─── Export Handler ───────────────────────────────────────────────
  const handleExport = () => {
    const wb = XLSX.utils.book_new();

    const addSheet = (data, name) => {
      const ws = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, name);
    };

    addSheet(guestHouseData, 'GuestHouse');
    addSheet(mealData, 'MealData');
    addSheet(manpowerFB, 'ManpowerFB');
    addSheet(vehicleData, 'Vehicles');
    addSheet(manpowerMaint, 'ManpowerMaint');
    addSheet(complaintsPie, 'Complaints');
    addSheet(complaintsStatus, 'ComplaintsStatus');
    addSheet(paintProgress, 'PaintProgress');

    XLSX.writeFile(wb, `Township_DPR_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // ─── Remove / Clear Data Handler ──────────────────────────────────
  const handleRemoveData = () => {
    setImportedData(null);
    setImportFileName('');
    setSelectedDate('');
    setShowComplaintsTable(false);
  };

  // ─── Visualization Renderer ────────────────────────────────────────
  const renderVisualization = () => {
    if (!importedData) {
      return (
        <motion.div
          key="upload-prompt"
          initial={{ opacity: 0, y: 15 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -15 }}
          transition={{ duration: 0.4, ease: [0.22, 1, 0.36, 1] }}
          className="md2-upload-placeholder"
        >
          <div className="md2-upload-card">
            <div className="md2-upload-icon-wrapper">
              <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" className="md2-upload-icon">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="17 8 12 3 7 8" />
                <line x1="12" y1="3" x2="12" y2="15" />
              </svg>
            </div>
            <h3 className="md2-upload-title">No DPR Data Loaded</h3>
            <p className="md2-upload-desc">
              Please upload the Township Daily Progress Report (DPR) spreadsheet to generate interactive charts, check occupancy rates, canteen meals, and maintenance metrics.
            </p>
            <button className="md2-upload-btn" onClick={() => fileInputRef.current?.click()}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '6px' }}>
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="17 8 12 3 7 8" />
                <line x1="12" y1="3" x2="12" y2="15" />
              </svg>
              <span>Select Excel File</span>
            </button>
            <div className="md2-upload-formats">
              Supported formats: <strong>.xlsx, .xls, .csv</strong>
            </div>
            <div className="md2-upload-requirements">
              <p className="md2-requirements-title">Expected Data Sheets:</p>
              <div className="md2-requirements-grid">
                <div className="md2-req-item">🏨 Guest House</div>
                <div className="md2-req-item">🍽️ Canteen Meals</div>
                <div className="md2-req-item">👷 F&B Manpower</div>
                <div className="md2-req-item">🚗 Fleet Movement</div>
                <div className="md2-req-item">🔧 Maintenance</div>
                <div className="md2-req-item">🎨 Paint Progress</div>
              </div>
            </div>
          </div>
        </motion.div>
      );
    }

    const contentVariants = {
      initial: { opacity: 0, scale: 0.98, y: 10 },
      animate: { opacity: 1, scale: 1, y: 0, transition: { duration: 0.4, ease: [0.22, 1, 0.36, 1] } },
      exit: { opacity: 0, scale: 0.98, y: -10, transition: { duration: 0.25 } },
    };

    switch (activeSection) {
      case 'guest-house':
        return (
          <motion.div key="guest-house" {...contentVariants} className="md2-viz-content md2-viz-container">
            <div className="md2-chart-grid">
              <div className="md2-chart-panel md2-chart-panel--large">
                <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Rooms Available vs Occupied</p>
                <ResponsiveContainer width="100%" height={300}>
                  <BarChart data={guestHouseData} barGap={6}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                    <XAxis dataKey="name" tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <YAxis tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <Tooltip {...TOOLTIP_STYLE} cursor={false} />
                    <Legend wrapperStyle={{ color: '#f8fafc', fontSize: 12, fontWeight: 'bold' }} />
                    <Bar dataKey="total" name="Total Rooms" fill="#0b4d8c" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="occupied" name="Occupied" fill="#c0293a" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="vacant" name="Vacant" fill="#2a9d8f" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="md2-chart-panel">
                <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Occupancy Rate</p>
                <ResponsiveContainer width="100%" height={300}>
                  <PieChart>
                    <Pie data={[
                      { name: 'Occupied', value: occupiedRooms },
                      { name: 'Vacant', value: Math.max(0, totalRooms - occupiedRooms) }
                    ]}
                      cx="50%" cy="50%" innerRadius={70} outerRadius={95}
                      dataKey="value" paddingAngle={4}>
                      <Cell fill="#c0293a" />
                      <Cell fill="#2a9d8f" />
                    </Pie>
                    <Tooltip {...TOOLTIP_STYLE} cursor={false} />
                    <Legend wrapperStyle={{ color: '#f8fafc', fontSize: 12, fontWeight: 'bold' }} />
                    <text x="50%" y="46%" textAnchor="middle" dominantBaseline="middle">
                      <tspan x="50%" dy="0" fill="#ffffff" style={{ fontSize: '24px', fontWeight: '800' }}>
                        {totalRooms > 0 ? Math.round((occupiedRooms / totalRooms) * 100) : 0}%
                      </tspan>
                      <tspan x="50%" dy="22" fill="#2a9d8f" style={{ fontSize: '11px', fontWeight: '700', letterSpacing: '0.8px', textTransform: 'uppercase' }}>
                        Occupied
                      </tspan>
                    </text>
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="md2-kpi-strip">
              {guestHouseData.map(g => (
                <div key={g.name} className="md2-kpi-mini" style={{ borderTop: '3px solid #0b4d8c' }}>
                  <span className="md2-kpi-mini-value" style={{ color: '#f8fafc' }}>{g.occupied}/{g.total}</span>
                  <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>{g.name}</span>
                </div>
              ))}
            </div>
          </motion.div>
        );

      case 'meal-report':
        return (
          <motion.div key="meal-report" {...contentVariants} className="md2-viz-content md2-viz-container">
            <div className="md2-chart-grid">
              <div className="md2-chart-panel md2-chart-panel--full">
                <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Meals by Location (Breakfast / Lunch / Dinner)</p>
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={mealData} barGap={4}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                    <XAxis dataKey="name" tick={{ fill: '#ffffff', fontSize: 11, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <YAxis tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <Tooltip {...TOOLTIP_STYLE} cursor={false} />
                    <Legend wrapperStyle={{ color: '#f8fafc', fontSize: 12, fontWeight: 'bold' }} />
                    <Bar dataKey="breakfast" name="Breakfast" fill="#0b4d8c" radius={[3, 3, 0, 0]} />
                    <Bar dataKey="lunch" name="Lunch" fill="#2a9d8f" radius={[3, 3, 0, 0]} />
                    <Bar dataKey="dinner" name="Dinner" fill="#c0293a" radius={[3, 3, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="md2-kpi-strip">
              {[
                { label: 'Breakfast', val: totalBreakfast, color: '#0b4d8c' },
                { label: 'Lunch', val: totalLunch, color: '#2a9d8f' },
                { label: 'Dinner', val: totalDinner, color: '#c0293a' },
                { label: 'Total Meals', val: totalMeals, color: '#f8fafc' },
              ].map(m => (
                <div key={m.label} className="md2-kpi-mini" style={{ borderLeft: `4px solid ${m.color}` }}>
                  <span className="md2-kpi-mini-value" style={{ color: m.color }}>{m.val.toLocaleString()}</span>
                  <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>{m.label}</span>
                </div>
              ))}
            </div>
          </motion.div>
        );

      case 'fb-manpower':
        return (
          <motion.div key="fb-manpower" {...contentVariants} className="md2-viz-content md2-viz-container">
            <div className="md2-chart-grid">
              <div className="md2-chart-panel md2-chart-panel--full">
                <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Actual vs Present by Category</p>
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={manpowerFB} layout="vertical" barGap={4}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                    <XAxis type="number" tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <YAxis dataKey="name" type="category" width={110} tick={{ fill: '#ffffff', fontSize: 11, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <Tooltip {...TOOLTIP_STYLE} cursor={false} />
                    <Legend wrapperStyle={{ color: '#f8fafc', fontSize: 12, fontWeight: 'bold' }} />
                    <Bar dataKey="actual" name="Actual" fill="#0b4d8c" radius={[0, 4, 4, 0]} />
                    <Bar dataKey="present" name="Present" fill="#2a9d8f" radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="md2-kpi-strip">
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #0b4d8c' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#0b4d8c' }}>{manpowerFB.reduce((s, m) => s + m.actual, 0)}</span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Total Actual</span>
              </div>
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #2a9d8f' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#2a9d8f' }}>{manpowerFB.reduce((s, m) => s + m.present, 0)}</span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Total Present</span>
              </div>
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #c0293a' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#c0293a' }}>
                  {manpowerFB.reduce((s, m) => s + m.actual, 0) > 0
                    ? Math.round((manpowerFB.reduce((s, m) => s + m.present, 0) / manpowerFB.reduce((s, m) => s + m.actual, 0)) * 100)
                    : 0}%
                </span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Attendance Rate</span>
              </div>
            </div>
          </motion.div>
        );

      case 'vehicle-fleet':
        return (
          <motion.div key="vehicle-fleet" {...contentVariants} className="md2-viz-content md2-viz-container">
            <div className="md2-chart-grid">
              <div className="md2-chart-panel md2-chart-panel--full">
                <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Distance Covered per Vehicle Today · Total: {totalKm} km</p>
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={vehicleData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                    <XAxis dataKey="name" tick={{ fill: '#ffffff', fontSize: 10, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <YAxis tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" unit=" km" />
                    <Tooltip {...TOOLTIP_STYLE} cursor={false} />
                    <Bar dataKey="km" name="Km Covered" radius={[4, 4, 0, 0]}>
                      {vehicleData.map((_, i) => (
                        <Cell key={i} fill={i % 2 === 0 ? '#2a9d8f' : '#0b4d8c'} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="md2-kpi-strip">
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #2a9d8f' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#2a9d8f' }}>{totalKm}</span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Total Km Today</span>
              </div>
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #0b4d8c' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#0b4d8c' }}>{vehicleData.length}</span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Vehicles Active</span>
              </div>
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #c0293a' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#c0293a' }}>
                  {vehicleData.length > 0 ? Math.round(totalKm / vehicleData.length) : 0}
                </span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Avg Km/Vehicle</span>
              </div>
            </div>
          </motion.div>
        );

      case 'complaints':
        return (
          <motion.div key="complaints" {...contentVariants} className="md2-viz-content md2-viz-container">
            <AnimatePresence mode="wait">
              {!showComplaintsTable ? (
                <motion.div
                  key="charts-view"
                  initial={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -80, transition: { duration: 0.35, ease: 'easeIn' } }}
                  style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: '20px' }}
                >
                  <div className="md2-chart-grid">
                    <div className="md2-chart-panel">
                      <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Jobs by Category</p>
                      <ResponsiveContainer width="100%" height={300}>
                        <PieChart>
                          <Pie data={complaintsPie} cx="50%" cy="50%"
                            outerRadius={95} dataKey="value" paddingAngle={3}
                            label={({ name, value }) => `${name}: ${value}`}
                            labelLine={{ stroke: 'rgba(255,255,255,0.4)', strokeWidth: 1.5 }}>
                            {complaintsPie.map((e, i) => <Cell key={i} fill={e.color || defaultComplaintsPie[i]?.color || '#0b4d8c'} />)}
                          </Pie>
                          <Tooltip {...TOOLTIP_STYLE} cursor={false} />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div className="md2-chart-panel md2-chart-panel--large">
                      <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Status by Category</p>
                      <ResponsiveContainer width="100%" height={300}>
                        <BarChart data={complaintsStatus} barGap={4}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                          <XAxis dataKey="name" tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                          <YAxis tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                          <Tooltip {...TOOLTIP_STYLE} cursor={false} />
                          <Legend wrapperStyle={{ color: '#f8fafc', fontSize: 12, fontWeight: 'bold' }} />
                          <Bar dataKey="completed" name="Completed" fill="#2a9d8f" radius={[4, 4, 0, 0]} />
                          <Bar dataKey="inProgress" name="In Progress" fill="#c0293a" radius={[4, 4, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                  <div className="md2-kpi-strip">
                    <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #f8fafc' }}>
                      <span className="md2-kpi-mini-value" style={{ color: '#f8fafc' }}>{totalComplaints}</span>
                      <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Total Jobs</span>
                    </div>
                    {complaintsPie.map((c, i) => (
                      <div key={c.name} className="md2-kpi-mini" style={{ borderLeft: `3px solid ${c.color || defaultComplaintsPie[i]?.color || '#0b4d8c'}` }}>
                        <span className="md2-kpi-mini-value" style={{ color: c.color || defaultComplaintsPie[i]?.color }}>{c.value}</span>
                        <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>{c.name}</span>
                      </div>
                    ))}
                  </div>

                  {/* Centered VIEW COMPLAINTS Button */}
                  <div style={{ display: 'flex', justifyContent: 'center', marginTop: '12px' }}>
                    <button
                      className="md2-view-complaints-btn"
                      onClick={() => setShowComplaintsTable(true)}
                      style={{
                        background: 'linear-gradient(135deg, #0b4d8c 0%, #1e40af 100%)',
                        color: '#ffffff',
                        border: 'none',
                        padding: '14px 32px',
                        borderRadius: '10px',
                        fontSize: '0.95rem',
                        fontWeight: '700',
                        letterSpacing: '0.5px',
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '10px',
                        boxShadow: '0 4px 20px rgba(11, 77, 140, 0.4)',
                        transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
                      }}
                    >
                      <span>VIEW COMPLAINTS</span>
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z" />
                        <circle cx="12" cy="12" r="3" />
                      </svg>
                    </button>
                  </div>
                </motion.div>
              ) : (
                <motion.div
                  key="table-view"
                  initial={{ opacity: 0, y: 80 }}
                  animate={{ opacity: 1, y: 0, transition: { duration: 0.4, ease: 'easeOut' } }}
                  exit={{ opacity: 0, y: 80, transition: { duration: 0.3 } }}
                  style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: '20px' }}
                >
                  {/* Complaints Card Panel */}
                  <div className="md2-chart-panel md2-chart-panel--full md2-complaints-card" style={{ padding: '24px' }}>
                    
                    {/* Complaints Header */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px', marginBottom: '20px' }}>
                      <div>
                        <h3 style={{ fontSize: '1.25rem', fontWeight: 800, color: '#f8fafc', margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#2a9d8f' }}>
                            <path d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z" />
                            <path d="m9 12 2 2 4-4" />
                          </svg>
                          Maintenance Complaints Ledger
                        </h3>
                        <p style={{ fontSize: '0.8rem', color: '#94a3b8', margin: '4px 0 0 0' }}>
                          Real-time complaint details extracted from the Daily Progress Report sheet.
                        </p>
                      </div>
                      
                      {/* Controls: Go back to Charts */}
                      <button
                        className="md2-back-charts-btn"
                        onClick={() => setShowComplaintsTable(false)}
                        style={{
                          background: 'rgba(255, 255, 255, 0.08)',
                          border: '1px solid rgba(255, 255, 255, 0.1)',
                          color: '#f8fafc',
                          padding: '10px 20px',
                          borderRadius: '8px',
                          fontSize: '0.85rem',
                          fontWeight: '600',
                          cursor: 'pointer',
                          display: 'flex',
                          alignItems: 'center',
                          gap: '6px',
                          transition: 'all 0.2s ease',
                        }}
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                          <polyline points="15 18 9 12 15 6" />
                        </svg>
                        <span>Back to Charts</span>
                      </button>
                    </div>

                    {/* Complaints Table */}
                    <div style={{ overflowX: 'auto', background: 'rgba(15, 23, 42, 0.3)', borderRadius: '10px', border: '1px solid rgba(255,255,255,0.05)' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem', textAlign: 'left' }}>
                        <thead>
                          <tr style={{ background: 'rgba(15, 23, 42, 0.6)', borderBottom: '1px solid rgba(255, 255, 255, 0.08)' }}>
                            <th style={{ padding: '16px 20px', fontWeight: '700', color: '#f1f5f9', width: '120px' }}>Department</th>
                            <th style={{ padding: '16px 20px', fontWeight: '700', color: '#f1f5f9', width: '100px' }}>Job No</th>
                            <th style={{ padding: '16px 20px', fontWeight: '700', color: '#f1f5f9' }}>Complaint Description</th>
                            <th style={{ padding: '16px 20px', fontWeight: '700', color: '#f1f5f9' }}>Completed Details</th>
                            <th style={{ padding: '16px 20px', fontWeight: '700', color: '#f1f5f9' }}>Pending Details</th>
                            <th style={{ padding: '16px 20px', fontWeight: '700', color: '#f1f5f9', width: '120px' }}>Feedback</th>
                          </tr>
                        </thead>
                        <tbody>
                          {complaintsList.length === 0 ? (
                            <tr>
                              <td colSpan={6} style={{ padding: '30px', textAlign: 'center', color: '#94a3b8' }}>
                                No complaints logged in this report.
                              </td>
                            </tr>
                          ) : (
                            complaintsList.map((comp, idx) => {
                              // Department specific colors
                              let deptColor = 'rgba(255,255,255,0.08)';
                              let deptBorder = 'rgba(255,255,255,0.15)';
                              let deptTextColor = '#ffffff';

                              if (comp.department === 'Civil') {
                                deptColor = 'rgba(192, 41, 58, 0.15)';
                                deptBorder = 'rgba(192, 41, 58, 0.3)';
                                deptTextColor = '#f87171';
                              } else if (comp.department === 'Plumbing') {
                                deptColor = 'rgba(11, 77, 140, 0.15)';
                                deptBorder = 'rgba(11, 77, 140, 0.3)';
                                deptTextColor = '#60a5fa';
                              } else if (comp.department === 'Seepage') {
                                deptColor = 'rgba(224, 122, 95, 0.15)';
                                deptBorder = 'rgba(224, 122, 95, 0.3)';
                                deptTextColor = '#fb923c';
                              } else if (comp.department === 'Carpentry') {
                                deptColor = 'rgba(42, 157, 143, 0.15)';
                                deptBorder = 'rgba(42, 157, 143, 0.3)';
                                deptTextColor = '#34d399';
                              } else if (comp.department === 'Electrical') {
                                deptColor = 'rgba(71, 85, 105, 0.15)';
                                deptBorder = 'rgba(71, 85, 105, 0.3)';
                                deptTextColor = '#cbd5e1';
                              }

                              return (
                                <tr
                                  key={idx}
                                  style={{
                                    borderBottom: '1px solid rgba(255, 255, 255, 0.04)',
                                    background: idx % 2 === 0 ? 'transparent' : 'rgba(255,255,255,0.01)',
                                    transition: 'background 0.2s ease',
                                  }}
                                  onMouseEnter={(e) => { e.currentTarget.style.background = 'rgba(255,255,255,0.03)'; }}
                                  onMouseLeave={(e) => { e.currentTarget.style.background = idx % 2 === 0 ? 'transparent' : 'rgba(255,255,255,0.01)'; }}
                                >
                                  {/* Department tag */}
                                  <td style={{ padding: '14px 20px' }}>
                                    <span
                                      style={{
                                        background: deptColor,
                                        border: `1px solid ${deptBorder}`,
                                        color: deptTextColor,
                                        padding: '4px 10px',
                                        borderRadius: '6px',
                                        fontSize: '0.75rem',
                                        fontWeight: '700',
                                        display: 'inline-block',
                                      }}
                                    >
                                      {comp.department || 'General'}
                                    </span>
                                  </td>
                                  
                                  {/* Job No */}
                                  <td style={{ padding: '14px 20px', color: '#f8fafc', fontWeight: '600' }}>
                                    {comp.jobNo || '—'}
                                  </td>
                                  
                                  {/* Description */}
                                  <td style={{ padding: '14px 20px', color: '#cbd5e1', lineHeight: '1.4' }}>
                                    {comp.description || '—'}
                                  </td>
                                  
                                  {/* Completed details */}
                                  <td style={{ padding: '14px 20px' }}>
                                    {comp.completedDetails ? (
                                      <span
                                        style={{
                                          color: '#34d399',
                                          background: 'rgba(52, 211, 153, 0.08)',
                                          border: '1px solid rgba(52, 211, 153, 0.2)',
                                          padding: '4px 10px',
                                          borderRadius: '6px',
                                          fontWeight: '600',
                                          fontSize: '0.8rem',
                                          display: 'inline-block',
                                        }}
                                      >
                                        ✓ {comp.completedDetails}
                                      </span>
                                    ) : (
                                      <span style={{ color: 'rgba(255,255,255,0.15)' }}>—</span>
                                    )}
                                  </td>
                                  
                                  {/* Pending details */}
                                  <td style={{ padding: '14px 20px' }}>
                                    {comp.pendingDetails ? (
                                      <span
                                        style={{
                                          color: '#fbbf24',
                                          background: 'rgba(251, 191, 36, 0.08)',
                                          border: '1px solid rgba(251, 191, 36, 0.2)',
                                          padding: '4px 10px',
                                          borderRadius: '6px',
                                          fontWeight: '600',
                                          fontSize: '0.8rem',
                                          display: 'inline-block',
                                        }}
                                      >
                                        ⚠ {comp.pendingDetails}
                                      </span>
                                    ) : (
                                      <span style={{ color: 'rgba(255,255,255,0.15)' }}>—</span>
                                    )}
                                  </td>
                                  
                                  {/* Feedback */}
                                  <td style={{ padding: '14px 20px', color: '#94a3b8' }}>
                                    {comp.feedback ? (
                                      <span
                                        style={{
                                          color: '#cbd5e1',
                                          background: 'rgba(203, 213, 225, 0.1)',
                                          padding: '2px 8px',
                                          borderRadius: '4px',
                                          fontSize: '0.75rem',
                                          fontWeight: '600',
                                        }}
                                      >
                                        {comp.feedback}
                                      </span>
                                    ) : (
                                      <span style={{ color: 'rgba(255,255,255,0.15)' }}>—</span>
                                    )}
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </motion.div>
        );

      case 'paint-work':
        return (
          <motion.div key="paint-work" {...contentVariants} className="md2-viz-content md2-viz-container">
            <div className="md2-chart-grid">
              <div className="md2-chart-panel md2-chart-panel--full">
                <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Target vs Achieved (Cumulative M²)</p>
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={paintProgress} barGap={6}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                    <XAxis dataKey="name" tick={{ fill: '#ffffff', fontSize: 11, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <YAxis tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <Tooltip {...TOOLTIP_STYLE} cursor={false} formatter={(v) => `${v.toLocaleString()} M²`} />
                    <Legend wrapperStyle={{ color: '#f8fafc', fontSize: 12, fontWeight: 'bold' }} />
                    <Bar dataKey="target" name="Target (M²)" fill="#0b4d8c" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="done" name="Achieved (M²)" fill="#2a9d8f" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="md2-kpi-strip">
              {paintProgress.map(p => (
                <div key={p.name} className="md2-kpi-mini" style={{ borderTop: `3px solid ${p.done >= p.target ? '#2a9d8f' : '#c0293a'}` }}>
                  <span className="md2-kpi-mini-value" style={{ color: p.done >= p.target ? '#2a9d8f' : '#c0293a' }}>
                    {p.target > 0 ? Math.round((p.done / p.target) * 100) : 0}%
                  </span>
                  <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>{p.name.replace('\n', ' ')}</span>
                </div>
              ))}
            </div>
          </motion.div>
        );

      case 'maint-manpower':
        return (
          <motion.div key="maint-manpower" {...contentVariants} className="md2-viz-content md2-viz-container">
            <div className="md2-chart-grid">
              <div className="md2-chart-panel md2-chart-panel--full">
                <p className="md2-chart-label" style={{ fontWeight: 'bold', color: '#f8fafc' }}>Actual vs Present</p>
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={manpowerMaint} layout="vertical" barGap={4}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                    <XAxis type="number" tick={{ fill: '#ffffff', fontSize: 12, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <YAxis dataKey="name" type="category" width={120} tick={{ fill: '#ffffff', fontSize: 11, fontWeight: 'bold' }} stroke="rgba(255,255,255,0.2)" />
                    <Tooltip {...TOOLTIP_STYLE} cursor={false} />
                    <Legend wrapperStyle={{ color: '#f8fafc', fontSize: 12, fontWeight: 'bold' }} />
                    <Bar dataKey="actual" name="Actual" fill="#0b4d8c" radius={[0, 4, 4, 0]} />
                    <Bar dataKey="present" name="Present" fill="#2a9d8f" radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="md2-kpi-strip">
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #0b4d8c' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#0b4d8c' }}>{manpowerMaint.reduce((s, m) => s + m.actual, 0)}</span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Total Actual</span>
              </div>
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #2a9d8f' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#2a9d8f' }}>{manpowerMaint.reduce((s, m) => s + m.present, 0)}</span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Total Present</span>
              </div>
              <div className="md2-kpi-mini" style={{ borderLeft: '4px solid #c0293a' }}>
                <span className="md2-kpi-mini-value" style={{ color: '#c0293a' }}>
                  {manpowerMaint.reduce((s, m) => s + m.actual, 0) > 0
                    ? Math.round((manpowerMaint.reduce((s, m) => s + m.present, 0) / manpowerMaint.reduce((s, m) => s + m.actual, 0)) * 100)
                    : 0}%
                </span>
                <span className="md2-kpi-mini-label" style={{ fontWeight: 'bold' }}>Attendance Rate</span>
              </div>
            </div>
          </motion.div>
        );

      default:
        return null;
    }
  };

  return (
    <div className="dashboard-page md2-page">
      {/* Hidden file input for import */}
      <input
        type="file"
        ref={fileInputRef}
        style={{ display: 'none' }}
        accept=".xlsx,.xls,.csv"
        onChange={handleFileSelect}
      />

      {/* ── Top Navigation Bar ── */}
      <nav className="md2-topbar">
        <div className="md2-topbar-inner">
          <div className="md2-topbar-left">
            <div className="nav-logo" onClick={() => navigate('/dashboard')} style={{ cursor: 'pointer' }}>
              <img src="/images/jsw_logo_clean.png" alt="JSW" className="nav-logo-img" />
            </div>
            <div className="md2-topbar-divider" />
            <h1 className="md2-topbar-title">Dashboard</h1>
          </div>

          {/* Centered HOME Button */}
          <div className="md2-topbar-center" style={{ position: 'absolute', left: '50%', transform: 'translateX(-50%)', display: 'flex', alignItems: 'center' }}>
            <button
              className="nav-link active md2-nav-home-btn"
              onClick={() => navigate('/dashboard')}
              style={{ background: 'rgba(255, 255, 255, 0.08)', color: '#ffffff', fontWeight: 600, border: 'none', cursor: 'pointer', padding: '8px 18px', borderRadius: '8px' }}
            >
              HOME
            </button>
          </div>

          <div className="md2-topbar-right">
            <div className="md2-search-box">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="11" cy="11" r="8" /><path d="m21 21-4.35-4.35" />
              </svg>
              <input type="text" placeholder="Search sections..." className="md2-search-input"
                onChange={(e) => {
                  const q = e.target.value.toLowerCase();
                  if (!q) return;
                  const match = SECTIONS.find(s => s.label.toLowerCase().includes(q));
                  if (match) setActiveSection(match.id);
                }}
              />
            </div>

            {/* Profile / Dropdown Toggle */}
            <div className="nav-profile" ref={profileMenuRef}>
              <button className="profile-btn" onClick={() => setProfileDropdownOpen(!profileDropdownOpen)}>
                <div className="profile-avatar">{user?.name?.charAt(0) || 'U'}</div>
                <div className="profile-info-compact" style={{ textAlign: 'left' }}>
                  <span className="profile-name" style={{ display: 'block', fontSize: '0.8rem', fontWeight: 600 }}>{user?.name || 'User'}</span>
                  <span className="profile-role-badge" style={{ display: 'block', fontSize: '0.65rem', color: '#2a9d8f', fontWeight: 'bold' }}>{user?.role || 'Staff'}</span>
                </div>
                <svg className={`dropdown-caret ${profileDropdownOpen ? 'open' : ''}`} width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" style={{ marginLeft: '4px', transition: 'transform 0.2s' }}>
                  <polyline points="6 9 12 15 18 9" />
                </svg>
              </button>

              <AnimatePresence>
                {profileDropdownOpen && (
                  <motion.div
                    className="profile-dropdown"
                    initial={{ opacity: 0, y: 8 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: 8 }}
                    transition={{ duration: 0.15 }}
                  >
                    <div className="dropdown-header">
                      <div className="dropdown-avatar">{user?.name?.charAt(0) || 'U'}</div>
                      <div style={{ textAlign: 'left' }}>
                        <div className="dropdown-name">{user?.name || 'User'}</div>
                        <div className="dropdown-role">{user?.role || 'Staff'}</div>
                      </div>
                    </div>
                    <div className="dropdown-divider" />
                    <button className="dropdown-item logout-item" onClick={() => { logout(); navigate('/login', { replace: true }); }}>
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}>
                        <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1-2-2h4" />
                        <polyline points="16 17 21 12 16 7" />
                        <line x1="21" y1="12" x2="9" y2="12" />
                      </svg>
                      <span>Logout</span>
                    </button>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </div>
        </div>
      </nav>

      {/* ── Sections Popup Overlay ── */}
      <AnimatePresence>
        {sectionsPanelOpen && (
          <>
            <motion.div
              className="md2-popup-backdrop"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              transition={{ duration: 0.25 }}
              onClick={() => setSectionsPanelOpen(false)}
            />
            <motion.div
              className="md2-popup-panel"
              initial={{ opacity: 0, y: -8, scale: 0.96 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: -8, scale: 0.96 }}
              transition={{ duration: 0.2, ease: [0.22, 1, 0.36, 1] }}
              style={{ transformOrigin: 'top left' }}
            >
              <div className="md2-popup-header">
                <h3 className="md2-popup-title">Sections</h3>
                <button className="md2-popup-close" onClick={() => setSectionsPanelOpen(false)}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
                  </svg>
                </button>
              </div>
              <div className="md2-popup-grid">
                {SECTIONS.map((section) => (
                  <button
                    key={section.id}
                    className={`md2-popup-item ${activeSection === section.id ? 'md2-popup-item--active' : ''}`}
                    onClick={() => { setActiveSection(section.id); setSectionsPanelOpen(false); }}
                  >
                    <span className="md2-popup-item-icon">{section.icon}</span>
                    <span className="md2-popup-item-label">{section.label}</span>
                    {activeSection === section.id && (
                      <span className="md2-popup-item-check">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                          <polyline points="20 6 9 17 4 12" />
                        </svg>
                      </span>
                    )}
                  </button>
                ))}
              </div>
              <div className="md2-popup-footer">
                <div className="md2-popup-kpi">
                  <span className="md2-popup-kpi-value">{totalMeals.toLocaleString()}</span>
                  <span className="md2-popup-kpi-label">Meals Today</span>
                </div>
                <div className="md2-popup-kpi-divider" />
                <div className="md2-popup-kpi">
                  <span className="md2-popup-kpi-value">{occupiedRooms}/{totalRooms}</span>
                  <span className="md2-popup-kpi-label">Rooms Occupied</span>
                </div>
                <div className="md2-popup-kpi-divider" />
                <div className="md2-popup-kpi">
                  <span className="md2-popup-kpi-value">{totalKm}</span>
                  <span className="md2-popup-kpi-label">Fleet Km</span>
                </div>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* ── Main Layout ── */}
      <div className="md2-layout">
        <main className="md2-content">
          {/* Content Header */}
          <div className="md2-content-header">
            <div className="md2-content-header-left">
              <span className="md2-content-icon">{currentSection?.icon}</span>
              <div>
                <AnimatePresence mode="wait">
                  {importedData ? (
                    <motion.h2
                      key="jsw-title"
                      initial={{ opacity: 0, y: -10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -10 }}
                      className="md2-content-title"
                      style={{ color: '#0b4d8c', fontWeight: 800 }}
                    >
                      JSW DAILY REPORT
                    </motion.h2>
                  ) : (
                    <motion.h2
                      key="generic-title"
                      initial={{ opacity: 0, y: -10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -10 }}
                      className="md2-content-title"
                    >
                      Township Daily Report
                    </motion.h2>
                  )}
                </AnimatePresence>
                <p className="md2-content-subtitle">
                  Data source: Township DPR Sheets
                  {importFileName && <span className="md2-import-badge"> · Imported: {importFileName}</span>}
                </p>
              </div>
            </div>
            <div className="md2-content-header-right">
              <span className="md2-date-badge">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '4px' }}>
                  <rect x="3" y="4" width="18" height="18" rx="2" ry="2" /><line x1="16" y1="2" x2="16" y2="6" /><line x1="8" y1="2" x2="8" y2="6" /><line x1="3" y1="10" x2="21" y2="10" />
                </svg>
                {selectedDate
                  ? new Date(selectedDate).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' })
                  : new Date().toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' })}
              </span>
            </div>
          </div>

          {/* ── Two-Panel Split ── */}
          <div className="md2-split">
            {/* Left Panel — Section List */}
            <div className="md2-split-left">
              <div className="md2-split-left-header">
                <span className="md2-split-left-label">Data Sections</span>
              </div>
              <div className="md2-split-left-list">
                {SECTIONS.map((section) => (
                  <button
                    key={section.id}
                    className={`md2-split-item ${activeSection === section.id ? 'md2-split-item--active' : ''}`}
                    onClick={() => setActiveSection(section.id)}
                  >
                    <span className="md2-split-item-icon">{section.icon}</span>
                    <span className="md2-split-item-label">{section.label}</span>
                    {activeSection === section.id && (
                      <motion.div className="md2-split-item-bar" layoutId="splitIndicator" />
                    )}
                  </button>
                ))}
              </div>
            </div>

            {/* Right Panel — Visualization */}
            <div className="md2-split-right">
              <div className="md2-split-right-header">
                <span className="md2-split-right-icon">{currentSection?.icon}</span>
                <span className="md2-split-right-title">{currentSection?.label}</span>

                {/* ── Calendar Picker Dropdown ── */}
                <div className="md2-date-picker-container" ref={dateDropdownRef}>
                  <button
                    type="button"
                    className={`md2-date-picker-trigger ${dateDropdownOpen ? 'active' : ''}`}
                    onClick={() => setDateDropdownOpen(!dateDropdownOpen)}
                  >
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                      <line x1="16" y1="2" x2="16" y2="6" />
                      <line x1="8" y1="2" x2="8" y2="6" />
                      <line x1="3" y1="10" x2="21" y2="10" />
                    </svg>
                    <span>{selectedDate ? selectedDate : 'Select Date'}</span>
                    <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" style={{ marginLeft: '2px', transform: dateDropdownOpen ? 'rotate(180deg)' : 'none', transition: 'transform 0.2s' }}>
                      <polyline points="6 9 12 15 18 9" />
                    </svg>
                  </button>

                  <AnimatePresence>
                    {dateDropdownOpen && (
                      <motion.div
                        className="md2-date-picker-menu"
                        initial={{ opacity: 0, y: 8 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, y: 8 }}
                        transition={{ duration: 0.15 }}
                      >
                        <div className="md2-date-picker-header">Historical Reports</div>
                        {availableDates.length === 0 ? (
                          <div style={{ padding: '8px 12px', fontSize: '0.75rem', color: 'var(--gray)' }}>No reports found</div>
                        ) : (
                          availableDates.map((d) => (
                            <button
                              key={d}
                              type="button"
                              className={`md2-date-picker-item ${selectedDate === d ? 'selected' : ''}`}
                              onClick={async () => {
                                setDateDropdownOpen(false);
                                await fetchReportForDate(d);
                              }}
                            >
                              <span>{d}</span>
                              {selectedDate === d && (
                                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#2a9d8f' }}>
                                  <polyline points="20 6 9 17 4 12" />
                                </svg>
                              )}
                            </button>
                          ))
                        )}
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              </div>
              <div className="md2-split-right-body">
                <AnimatePresence mode="wait">
                  {renderVisualization()}
                </AnimatePresence>
              </div>

              {/* Holographic Loader Overlay */}
              <AnimatePresence>
                {isAnalyzing && (
                  <motion.div
                    className="md2-analysis-overlay"
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    exit={{ opacity: 0 }}
                    transition={{ duration: 0.3 }}
                  >
                    <div className="md2-hologram-container">
                      <div className="md2-holo-spinner" />
                      <div className="md2-holo-radar" />
                      <div className="md2-holo-scanner">
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" className="md2-holo-icon">
                          <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" />
                          <polyline points="14 2 14 8 20 8" />
                          <path d="M8 13h8" />
                          <path d="M8 17h8" />
                          <path d="M8 9h2" />
                        </svg>
                        <div className="md2-scanner-sweep" />
                      </div>
                    </div>
                    <h3 className="md2-analysis-title">Analyzing Spreadsheet</h3>
                    <p className="md2-analysis-subtitle">
                      JSW Group AI Data Engine is running optimized parsing algorithms on your Daily Progress Report.
                    </p>
                    <div className="md2-analysis-terminal">
                      <span className="md2-terminal-prompt">SYS_KERN:&gt;</span>
                      <span className="md2-terminal-text" key={analyzingStep}>
                        {analyzingStep}
                      </span>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </div>

          {/* Bottom Bar */}
          <div className="md2-bottom-bar">
            <div className="md2-bottom-bar-left">
              <span className="md2-section-badge">
                {currentSection?.icon} {currentSection?.label}
              </span>
              <span className="md2-data-info">
                {user?.name && `· ${user.name}`} · JSW Group · Township Management Portal
              </span>
            </div>
            {/* Bottom Actions */}
            <div style={{ display: 'flex', gap: '12px' }}>
              {importedData && (
                <button className="md2-action-btn md2-action-btn--remove" onClick={handleRemoveData} title="Remove Ingested Data">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '6px' }}>
                    <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                    <line x1="9" y1="9" x2="15" y2="15" />
                    <line x1="15" y1="9" x2="9" y2="15" />
                  </svg>
                  <span>Remove Data</span>
                </button>
              )}
              <button className="md2-action-btn md2-action-btn--export" onClick={handleExport} title="Export to Excel">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '6px' }}>
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="7 10 12 15 17 10" />
                  <line x1="12" y1="15" x2="12" y2="3" />
                </svg>
                <span>Export Data</span>
              </button>
            </div>
          </div>

        </main>
      </div>

      {/* ── Date Ingestion Confirmation Modal ── */}
      <AnimatePresence>
        {showDateModal && (
          <motion.div
            className="md2-date-modal-overlay"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
          >
            <motion.div
              className="md2-date-modal"
              initial={{ scale: 0.9, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.9, y: 20 }}
              transition={{ type: 'spring', damping: 25, stiffness: 350 }}
            >
              <div className="md2-date-modal-icon">📅</div>
              <h3 className="md2-date-modal-title">Confirm Report Date</h3>
              <p className="md2-date-modal-desc">
                Specify the operation date for the Township Daily Progress Report. This ensures data records are tracked accurately in historical audits.
              </p>
              <div className="md2-date-input-group">
                <label className="md2-date-input-label">Report Date</label>
                <input
                  type="date"
                  className="md2-date-input"
                  value={inputDate}
                  onChange={(e) => setInputDate(e.target.value)}
                  required
                />
              </div>
              <div className="md2-date-modal-actions">
                <button
                  type="button"
                  className="md2-date-modal-btn md2-date-modal-btn--cancel"
                  onClick={() => {
                    setShowDateModal(false);
                    setPendingFile(null);
                  }}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  className="md2-date-modal-btn md2-date-modal-btn--confirm"
                  onClick={handleImportConfirm}
                  disabled={!inputDate}
                >
                  Confirm & Ingest
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
