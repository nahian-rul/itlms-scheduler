import React, { useState, useEffect, useMemo, useRef } from 'react';
import { ErrorBoundary } from './ErrorBoundary';
import {
    ChevronRight,
    Calendar,
    Plus,
    ArrowLeft,
    Printer,
    Copy,
    Trash2,
    Edit2,
    Pencil,
    Save,
    X,
    ChevronDown,
    MoreVertical,
    Combine,
    Split,
    UserPlus,
    MinusCircle,
    FileText,
    Info
} from 'lucide-react';

const formatTime12 = (time24) => {
    if (!time24 || typeof time24 !== 'string') return '';
    const [h, m] = time24.split(':');
    let hours = parseInt(h, 10);
    const suffix = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12 || 12;
    return `${hours}:${m} ${suffix}`;
};

// --- Mock Data Constants ---

const INITIAL_COURSES = [
    { id: 'C1', code: 'FTC-207', name: '207th Foundation Training Course', type: 'Long Term', status: 'Published' },
    { id: 'C2', code: 'EAM-137', name: '137th EAM Course for Head of Secondary Level Institution', type: 'Short Term', status: 'Draft' },
];

const INITIAL_BATCHES = [
    { id: 'B1', courseId: 'C1', name: 'Batch 207', startDate: '2026-03-02', endDate: '2026-06-29', trainees: 80, status: 'Ongoing' },
    { id: 'B2', courseId: 'C2', name: 'Batch 137', startDate: '2026-03-31', endDate: '2026-04-20', trainees: 45, status: 'Upcoming' },
];

const MASTER_SECTIONS = ['Padma', 'Meghna', 'Jamuna', 'Karnaphuli', 'Surma', 'Rupsha', 'Buriganga', 'Arial Kha'];

const MOCK_MODULES = [
    { id: 'M1', code: '1.01', name: 'Public Administration' },
    { id: 'M2', code: '2.04', name: 'ICT for Managers' },
    { id: 'M3', code: '5.08', name: 'Service Rules & Financial Management' },
    { id: 'M4', code: '3.02', name: 'Leadership & Change Management' },
    { id: 'M5', code: '4.05', name: 'Development Economics' },
    { id: 'M6', code: '6.11', name: 'Environmental Management' },
    { id: 'M7', code: '7.03', name: 'Ethics & Anti-Corruption' },
    { id: 'M8', code: '8.09', name: 'Communication & Presentation Skills' }
];

const MOCK_CONTENT = {
    'M1': ['Citizen Charter', 'Governance Models', 'Policy Making', 'Decentralization & Local Government', 'E-Governance Frameworks'],
    'M2': ['Excel Advanced', 'Cyber Security', 'AI in Office', 'Data Visualization & Power BI', 'Cloud Computing Basics', 'Digital Transformation Strategy'],
    'M3': ['Audit Procedures', 'GFR & PPA', 'Pension Rules', 'Budget Preparation & Execution', 'Public Procurement Rules 2008'],
    'M4': ['Transformational Leadership', 'Strategic Change Management', 'Team Building & Motivation', 'Conflict Resolution', 'Organizational Behavior', 'Decision Making Under Uncertainty'],
    'M5': ['SDG Implementation Framework', 'Poverty Reduction Strategies', 'Foreign Aid & Development Finance', 'Inclusive Growth Policies', 'Public-Private Partnership Models'],
    'M6': ['Climate Change Adaptation', 'Environmental Impact Assessment', 'Green Public Procurement', 'Disaster Risk Reduction', 'Sustainable Development Goals'],
    'M7': ['Code of Conduct for Civil Servants', 'Anti-Corruption Laws & Frameworks', 'Integrity in Public Service', 'Whistle-Blower Protection', 'Asset Disclosure & Transparency'],
    'M8': ['Effective Report Writing', 'Public Speaking & Presentation', 'Note Writing & Office Communication', 'Media Engagement Skills', 'Negotiation Techniques']
};

const MOCK_SPEAKERS = [
    'Mr. Mustafizur Rahman',
    'Ms. Tamanna Islam',
    'Dr. Abdullah Al Mamun',
    'Prof. Shoukat Ali',
    'Dr. Nargis Begum',
    'Mr. Rezaul Karim',
    'Ms. Farhana Hoque',
    'Prof. Monirul Islam',
    'Dr. Shamim Ara',
    'Mr. Khairul Bashar',
    'Ms. Dilruba Yasmin',
    'Dr. Ataur Rahman'
];

const MOCK_ACTIVITIES = [
    'Morning PT/Yoga',
    'Car Driving Training',
    'Library Work',
    'Games & Sports',
    'Study Tour Briefing',
    'Tea Break',
    'Prayer & Lunch',
    'Cultural Evening Program',
    'Group Discussion & Syndicate Work',
    'Film Show & Documentary',
    'National Anthem & Flag Ceremony',
    'Indoor Games',
    'Meditation & Mindfulness',
    'Garden Walk & Nature Tour',
    'Graduation Rehearsal'
];

// --- Utility UI Components ---

const Card = ({ children, className = "" }) => (
    <div className={`bg-white rounded-lg border border-slate-200 shadow-sm ${className}`}>
        {children}
    </div>
);

const Button = ({ children, onClick, variant = "primary", className = "", icon: Icon, disabled = false }) => {
    const variants = {
        primary: "bg-blue-600 text-white hover:bg-blue-700 disabled:bg-blue-300",
        secondary: "bg-slate-100 text-slate-700 hover:bg-slate-200",
        outline: "border border-slate-300 text-slate-700 hover:bg-slate-50",
        danger: "bg-red-50 text-red-600 hover:bg-red-100 border border-red-200",
        ghost: "text-slate-500 hover:bg-slate-100"
    };
    return (
        <button
            disabled={disabled}
            onClick={onClick}
            className={`px-4 py-2 rounded-md font-medium text-sm flex items-center justify-center gap-2 transition-colors ${variants[variant]} ${className}`}
        >
            {Icon && <Icon size={16} />}
            {children}
        </button>
    );
};

// --- Main Application Component ---

export default function App() {
    const [view, setView] = useState('courses');
    const [selectedCourse, setSelectedCourse] = useState(null);
    const [selectedBatch, setSelectedBatch] = useState(null);
    const [selectedDate, setSelectedDate] = useState('2026-03-04');
    const [schedules, setSchedules] = useState({});
    const [cloneModal, setCloneModal] = useState(false);
    const [guestSpeakers, setGuestSpeakers] = useState([
        { id: 'GS1', name: 'John Doe', designation: 'Visiting Professor', email: 'john@example.com', phone: '01712345678' }
    ]);

    // Navigation handlers
    const openBatches = (course) => {
        setSelectedCourse(course);
        setView('batches');
    };

    const openBatchDetails = (batch) => {
        setSelectedBatch(batch);
        setView('details');
    };

    const openBuilder = (date) => {
        console.log("[App.jsx] openBuilder called with date:", date);
        const finalDate = date || '2026-03-04';
        console.log("[App.jsx] setting selectedDate:", finalDate);
        setSelectedDate(finalDate);
        setView('builder');
    };

    const openExport = () => setView('export');

    const goBack = () => {
        if (view === 'export') setView('builder');
        else if (view === 'builder') setView('details');
        else if (view === 'details') setView('batches');
        else if (view === 'batches') setView('courses');
    };

    const saveSchedule = (date, scheduleData) => {
        setSchedules(prev => ({
            ...prev,
            [date]: { ...scheduleData, saved: true, savedAt: new Date().toISOString() }
        }));
        setView('details');
    };

    const cloneSchedule = (fromDate, toDate) => {
        if (!schedules[fromDate] || !toDate) return;
        setSchedules(prev => ({
            ...prev,
            [toDate]: { ...schedules[fromDate], saved: false, savedAt: null }
        }));
        setCloneModal(false);
        openBuilder(toDate);
    };

    return (
        <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
            <header className="bg-white border-b border-slate-200 h-16 flex items-center px-6 sticky top-0 z-50 shadow-sm">
                <div className="flex items-center gap-3">
                    <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold">N</div>
                    <h1 className="font-bold text-lg tracking-tight">NAEM Training Portal</h1>
                </div>

                {view !== 'courses' && (
                    <div className="flex items-center ml-8 text-sm text-slate-500 overflow-hidden whitespace-nowrap">
                        <button onClick={() => setView('courses')} className="hover:text-blue-600">Courses</button>
                        <ChevronRight size={14} className="mx-2" />
                        {selectedCourse && (
                            <>
                                <button onClick={() => setView('batches')} className="hover:text-blue-600 truncate max-w-[150px]">{selectedCourse.name}</button>
                                <ChevronRight size={14} className="mx-2" />
                            </>
                        )}
                        {selectedBatch && (
                            <span className="text-slate-900 font-medium">{selectedBatch.name}</span>
                        )}
                    </div>
                )}
            </header>

            <main className={view === 'builder' ? "w-full" : "p-6 max-w-[1600px] mx-auto"} style={view === 'builder' ? { height: 'calc(100vh - 64px)' } : {}}>
                {view === 'courses' && <CourseList onSelect={openBatches} />}
                {view === 'batches' && <BatchList course={selectedCourse} onSelect={openBatchDetails} />}
                {view === 'details' && <BatchDetails batch={selectedBatch} course={selectedCourse} onOpenBuilder={openBuilder} schedules={schedules} onClone={() => setCloneModal(true)} />}
                {view === 'builder' && (
                    <ErrorBoundary>
                        <ScheduleBuilder
                            batch={selectedBatch}
                            course={selectedCourse}
                            date={selectedDate}
                            onBack={goBack}
                            onExport={openExport}
                            schedules={schedules}
                            setSchedules={setSchedules}
                            onSaveSchedule={saveSchedule}
                            onClone={() => setCloneModal(true)}
                            guestSpeakers={guestSpeakers}
                            setGuestSpeakers={setGuestSpeakers}
                        />
                    </ErrorBoundary>
                )}
                {view === 'export' && (
                    <ScheduleExport
                        batch={selectedBatch}
                        course={selectedCourse}
                        date={selectedDate}
                        schedule={schedules[selectedDate]}
                        onBack={goBack}
                    />
                )}
            </main>
            {cloneModal && (
                <CloneModal
                    schedules={schedules}
                    onClose={() => setCloneModal(false)}
                    onClone={cloneSchedule}
                />
            )}
        </div>
    );
}

// --- List Views ---

function CourseList({ onSelect }) {
    return (
        <div className="animate-in fade-in slide-in-from-bottom-2 duration-300">
            <div className="flex justify-between items-center mb-6">
                <h2 className="text-2xl font-bold">Course Management</h2>
                <Button icon={Plus}>Add Course</Button>
            </div>
            <Card>
                <table className="w-full text-left">
                    <thead className="bg-slate-50 border-b border-slate-200">
                        <tr>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">SL</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Code</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Course Name</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Type</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Status</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase text-right">Action</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                        {INITIAL_COURSES.map((course, idx) => (
                            <tr key={course.id} className="hover:bg-slate-50 transition-colors">
                                <td className="px-6 py-4 text-sm">{idx + 1}</td>
                                <td className="px-6 py-4 text-sm font-medium">{course.code}</td>
                                <td className="px-6 py-4 text-sm">{course.name}</td>
                                <td className="px-6 py-4 text-sm">{course.type}</td>
                                <td className="px-6 py-4 text-sm">
                                    <span className={`px-2 py-1 rounded-full text-[10px] font-bold uppercase ${course.status === 'Published' ? 'bg-green-100 text-green-700' : 'bg-amber-100 text-amber-700'}`}>
                                        {course.status}
                                    </span>
                                </td>
                                <td className="px-6 py-4 text-sm text-right">
                                    <Button variant="secondary" onClick={() => onSelect(course)}>Batch</Button>
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </Card>
        </div>
    );
}

function BatchList({ course, onSelect }) {
    return (
        <div className="animate-in fade-in slide-in-from-bottom-2 duration-300">
            <div className="flex justify-between items-center mb-6">
                <h2 className="text-2xl font-bold">Batches for {course.code}</h2>
                <Button icon={Plus}>Create Batch</Button>
            </div>
            <Card>
                <table className="w-full text-left">
                    <thead className="bg-slate-50 border-b border-slate-200">
                        <tr>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">SL</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Batch Name</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Timeline</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase text-center">Trainees</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Status</th>
                            <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase text-right">Action</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                        {INITIAL_BATCHES.filter(b => b.courseId === course.id).map((batch, idx) => (
                            <tr key={batch.id} className="hover:bg-slate-50 transition-colors">
                                <td className="px-6 py-4 text-sm">{idx + 1}</td>
                                <td className="px-6 py-4 text-sm font-semibold">{batch.name}</td>
                                <td className="px-6 py-4 text-sm text-slate-500">{batch.startDate} – {batch.endDate}</td>
                                <td className="px-6 py-4 text-sm text-center">{batch.trainees}</td>
                                <td className="px-6 py-4 text-sm">
                                    <span className="px-2 py-1 rounded-full text-[10px] font-bold uppercase bg-blue-100 text-blue-700">
                                        {batch.status}
                                    </span>
                                </td>
                                <td className="px-6 py-4 text-sm text-right">
                                    <Button variant="outline" onClick={() => onSelect(batch)}>View</Button>
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </Card>
        </div>
    );
}

function BatchDetails({ batch, course, onOpenBuilder, schedules, onClone }) {
    const [activeTab, setActiveTab] = useState('Schedule');
    const tabs = ['Overview', 'Trainees', 'Modules', 'Schedule', 'Evaluation', 'Certificate'];

    const savedSchedules = Object.entries(schedules)
        .filter(([, s]) => s.saved)
        .sort(([a], [b]) => new Date(a) - new Date(b));

    const fmtDate = (d) => d.split('-').reverse().join('.');
    const dayName = (d) => ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][new Date(d).getDay()];

    return (
        <div className="space-y-6 animate-in fade-in duration-300">
            <div>
                <h2 className="text-2xl font-bold mb-1">{course.name}</h2>
                <p className="text-slate-500">{batch.name} | {batch.startDate} to {batch.endDate}</p>
            </div>

            <div className="border-b border-slate-200">
                <nav className="flex gap-8">
                    {tabs.map(tab => (
                        <button key={tab} onClick={() => setActiveTab(tab)}
                            className={`pb-3 text-sm font-medium relative ${activeTab === tab ? 'text-blue-600' : 'text-slate-400 hover:text-slate-600'}`}>
                            {tab}
                            {activeTab === tab && <div style={{position:'absolute',bottom:0,left:0,right:0,height:'2px',background:'#2563eb',borderRadius:'9999px'}} />}
                        </button>
                    ))}
                </nav>
            </div>

            {activeTab === 'Schedule' && (
                <div className="space-y-4">
                    <div className="flex justify-between items-center">
                        <div className="flex gap-3 items-center">
                            <p className="text-sm text-slate-500">{savedSchedules.length} schedule{savedSchedules.length !== 1 ? 's' : ''} saved</p>
                            <div className="h-4 w-px bg-slate-200 mx-1" />
                            <div className="flex items-center gap-2">
                                <label className="text-xs font-bold text-slate-400 uppercase">New Schedule Date:</label>
                                <input 
                                    type="date" 
                                    id="new-schedule-date"
                                    className="px-3 py-1.5 border border-slate-200 rounded-md text-sm outline-none focus:ring-2 focus:ring-blue-500" 
                                    defaultValue={new Date().toISOString().split('T')[0]} 
                                />
                            </div>
                        </div>
                        <div className="flex gap-2">
                            <Button variant="outline" icon={Copy} onClick={onClone}>Clone Schedule</Button>
                            <Button onClick={() => {
                                const dateInput = document.getElementById('new-schedule-date');
                                onOpenBuilder(dateInput?.value);
                            }} icon={Plus}>Create Schedule</Button>
                        </div>
                    </div>

                    <Card>
                        <table className="w-full text-left">
                            <thead className="bg-slate-50 border-b border-slate-200">
                                <tr>
                                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">SL</th>
                                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Date</th>
                                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Day</th>
                                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Sessions</th>
                                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase">Status</th>
                                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase text-right">Action</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                                {savedSchedules.length === 0 ? (
                                    <tr><td colSpan={6} style={{padding:'40px',textAlign:'center',color:'#94a3b8',fontSize:'14px'}}>No schedules saved yet. Create and save a schedule to see it here.</td></tr>
                                ) : savedSchedules.map(([date, sched], idx) => (
                                    <tr key={date} className="hover:bg-slate-50 transition-colors">
                                        <td className="px-6 py-4 text-sm">{idx + 1}</td>
                                        <td className="px-6 py-4 text-sm font-medium text-blue-600">{fmtDate(date)}</td>
                                        <td className="px-6 py-4 text-sm text-slate-500">{dayName(date)}</td>
                                        <td className="px-6 py-4 text-sm text-slate-500">{sched.columns?.length || 0} columns · {sched.sections?.length || 0} sections</td>
                                        <td className="px-6 py-4 text-sm">
                                            <span style={{padding:'2px 10px',borderRadius:'99px',fontSize:'10px',fontWeight:700,textTransform:'uppercase',background:'#dcfce7',color:'#15803d'}}>Saved</span>
                                        </td>
                                        <td className="px-6 py-4 text-sm text-right space-x-2">
                                            <Button variant="ghost" className="p-1 min-w-0" icon={Edit2} onClick={() => onOpenBuilder(date)} />
                                            <Button variant="ghost" className="p-1 min-w-0" icon={Printer} onClick={() => onOpenBuilder(date)} />
                                            <Button variant="ghost" className="p-1 min-w-0 text-red-500" icon={Trash2} />
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </Card>
                </div>
            )}

            {activeTab !== 'Schedule' && (
                <div className="flex items-center justify-center py-20 text-slate-400">
                    <div className="text-center">
                        <FileText size={48} className="mx-auto mb-4 opacity-30" />
                        <p className="font-medium">No data yet for the <strong>{activeTab}</strong> tab.</p>
                    </div>
                </div>
            )}
        </div>
    );
}


// --- Schedule Builder Module ---

const ScheduleBuilder = ({ batch, course, date, onBack, onExport, schedules, setSchedules, onSaveSchedule, onClone, guestSpeakers, setGuestSpeakers }) => {
    console.log("[ScheduleBuilder] Mounting with date:", date, "batch:", batch);
    const [editingCell, setEditingCell] = useState(null);
    const [editingHeader, setEditingHeader] = useState(null);
    const [isAddingSection, setIsAddingSection] = useState(false);
    const [isManagingGuidelines, setIsManagingGuidelines] = useState(false);
    const [mergeConflict, setMergeConflict] = useState(null); // { section, colId, data, conflictSection }
    const [saved, setSaved] = useState(schedules[date]?.saved || false);
    const [selectedHeaders, setSelectedHeaders] = useState(new Set());
    const [mergingHeaders, setMergingHeaders] = useState(null); // { id?: string, title: string, colIds: [] }
    const [splittingHeader, setSplittingHeader] = useState(null); // { mergeId, colIds, currentNames, defaultNames }
    const [splittingSession, setSplittingSession] = useState(null); // { section, col, cell }
    const [selectedCells, setSelectedCells] = useState(new Set()); // Set of "section_colId"

    const handleSplitSession = (colId, splitPoint) => {
        const colIdx = columns.findIndex(c => c.id === colId);
        const col = columns[colIdx];
        
        const newCols = [...columns];
        const newCol1 = { ...col, endTime: splitPoint, title: `${col.title} (Part 1)` };
        const newCol2 = { id: `col_split_${Date.now()}`, title: `${col.title} (Part 2)`, startTime: splitPoint, endTime: col.endTime };
        
        newCols.splice(colIdx, 1, newCol1, newCol2);
        
        // Duplicate data for the new column for all sections to maintain "merged" look initially but allow edits
        const newCells = { ...cells };
        sections.forEach(s => {
            const originalCell = cells[`${s}_${colId}`];
            if (originalCell) {
                newCells[`${s}_${newCol2.id}`] = { ...originalCell };
            }
        });
        
        updateSchedule({ columns: newCols, cells: newCells });
    };

    const isSaved = schedules[date]?.saved;

    const handleSave = () => {
        onSaveSchedule(date, currentSchedule);
        setSaved(true);
    };

    // Core state structure
    const currentSchedule = useMemo(() => schedules[date] || {
        date,
        sections: ['Padma', 'Meghna', 'Jamuna', 'Karnaphuli'],
        columns: [
            { id: 'col1', title: 'Morning Activities', startTime: '06:00', endTime: '08:30' },
            { id: 'col2', title: 'Session 1',          startTime: '09:00', endTime: '10:00' },
            { id: 'col3', title: 'Session 2',          startTime: '10:15', endTime: '11:15' },
            { id: 'col4', title: 'Session 3',          startTime: '11:30', endTime: '12:30' },
            { id: 'col5', title: 'Lunch & Prayer',     startTime: '12:30', endTime: '14:00' },
            { id: 'col6', title: 'Session 4',          startTime: '14:00', endTime: '15:00' },
            { id: 'col7', title: 'Session 5',          startTime: '15:15', endTime: '16:15' },
            { id: 'col8', title: 'Evening Activities', startTime: '16:30', endTime: '18:00' },
        ],
        cells: {},
        guidelines: [
            { id: 'g1', name: 'Dress Code',   instruction: 'White Shirt, Dark Pant & Blue Tie (Male). Standard Formal Dress (Female).' },
            { id: 'g2', name: 'Punctuality',  instruction: 'All participants must be present 5 minutes before each session starts.' },
            { id: 'g3', name: 'Mobile Phone', instruction: 'Mobile phones must be on silent mode during all sessions.' }
        ],
        headerMerges: []
    }, [schedules, date]);

    const { columns, cells, sections, guidelines, headerMerges = [] } = currentSchedule;

    const dayNum = useMemo(() => {
        const start = new Date(batch.startDate);
        const curr = new Date(date);
        return Math.ceil(Math.abs(curr - start) / (1000 * 60 * 60 * 24)) + 1;
    }, [batch.startDate, date]);

    const toggleHeaderSelection = (colId) => {
        const newSelected = new Set(selectedHeaders);
        if (newSelected.has(colId)) {
            newSelected.delete(colId);
        } else {
            // Check if sequential
            if (newSelected.size > 0) {
                const colIdx = columns.findIndex(c => c.id === colId);
                const sortedIndices = Array.from(newSelected).map(id => columns.findIndex(c => c.id === id)).sort((a, b) => a - b);
                const min = sortedIndices[0];
                const max = sortedIndices[sortedIndices.length - 1];
                
                if (colIdx !== min - 1 && colIdx !== max + 1) {
                    alert("Please select sequential columns to merge.");
                    return;
                }
            }
            newSelected.add(colId);
        }
        setSelectedHeaders(newSelected);
    };

    const handleMergeHeaders = () => {
        const sortedCols = Array.from(selectedHeaders).sort((a, b) => {
            const idxA = columns.findIndex(c => c.id === a);
            const idxB = columns.findIndex(c => c.id === b);
            return idxA - idxB;
        });
        const colNames = sortedCols.map(id => columns.find(c => c.id === id)?.title || '');
        const autoTitle = colNames.join(' + ');
        setMergingHeaders({ title: autoTitle, colIds: sortedCols });
    };

    const confirmMerge = (title) => {
        if (mergingHeaders.id) {
            updateSchedule({
                headerMerges: headerMerges.map(m => m.id === mergingHeaders.id ? { ...m, title } : m)
            });
        } else {
            updateSchedule({
                headerMerges: [...headerMerges, { id: `merge_${Date.now()}`, title, colIds: mergingHeaders.colIds }]
            });
        }
        setMergingHeaders(null);
        setSelectedHeaders(new Set());
    };

    const handleUnmerge = (mergeId) => {
        const merge = headerMerges.find(m => m.id === mergeId);
        if (!merge) return;

        let splitNames = [];
        if (merge.title.includes(' + ')) {
            splitNames = merge.title.split(' + ');
            // Ensure we have enough parts for all columns
            while (splitNames.length < merge.colIds.length) splitNames.push('');
        } else {
            // Single name logic: 1st column gets the name, others empty
            splitNames = [merge.title, ...Array(merge.colIds.length - 1).fill('')];
        }

        const defaultNames = merge.colIds.map(id => columns.find(c => c.id === id)?.title || '');

        setSplittingHeader({
            mergeId,
            colIds: merge.colIds,
            currentNames: splitNames,
            defaultNames
        });
    };

    const confirmSplit = (newNames) => {
        const updatedColumns = columns.map(col => {
            const splitIdx = splittingHeader.colIds.indexOf(col.id);
            if (splitIdx !== -1) {
                return { ...col, title: newNames[splitIdx] || `Col ${splitIdx + 1}` };
            }
            return col;
        });

        updateSchedule({
            columns: updatedColumns,
            headerMerges: headerMerges.filter(m => m.id !== splittingHeader.mergeId)
        });
        setSplittingHeader(null);
    };

    const removeHeaderMerge = (id) => {
        handleUnmerge(id);
    };

    const updateSchedule = (newData) => {
        const newSchedules = { ...schedules };
        newSchedules[date] = { ...currentSchedule, ...newData };
        setSchedules(newSchedules);
    };

    const addColumn = () => {
        updateSchedule({
            columns: [...currentSchedule.columns, { id: `col_${Date.now()}`, title: `Session ${currentSchedule.columns.length + 1}`, startTime: '09:00', endTime: '10:00' }]
        });
    };

    const addColumnAfter = (id) => {
        const colIdx = currentSchedule.columns.findIndex(c => c.id === id);
        if (colIdx === -1) return;
        
        const newCol = { id: `col_${Date.now()}`, title: `New Session`, startTime: '09:00', endTime: '10:00' };
        const newCols = [...currentSchedule.columns];
        newCols.splice(colIdx + 1, 0, newCol);
        
        updateSchedule({ columns: newCols });
    };

    const updateHeader = (id, data) => {
        updateSchedule({
            columns: columns.map(c => c.id === id ? { ...c, ...data } : c)
        });
    };

    const addSection = (name) => {
        if (!name) return;
        updateSchedule({ sections: [...sections, name] });
        setIsAddingSection(false);
    };

    const removeSection = (name) => {
        updateSchedule({ sections: sections.filter(s => s !== name) });
    };

    const handleBulkMergeSessions = () => {
        const cellKeys = Array.from(selectedCells);
        if (cellKeys.length < 2) return;

        const masterKey = cellKeys[0];
        const [mSection, mColId] = masterKey.split('_');
        const masterCell = currentSchedule.cells[masterKey];
        
        if (!masterCell) {
            alert("The first selected cell must have content to be the master.");
            return;
        }

        const sectionsToMerge = cellKeys.slice(1).map(k => k.split('_')[0]);
        
        // Update master to include these sections
        saveCell(mSection, mColId, {
            ...masterCell,
            isMerging: true,
            mergedSections: [...(masterCell.mergedSections || []), ...sectionsToMerge]
        });
        setSelectedCells(new Set());
    };

    const toggleCellSelection = (key) => {
        const newSelected = new Set(selectedCells);
        if (newSelected.has(key)) newSelected.delete(key);
        else newSelected.add(key);
        setSelectedCells(newSelected);
    };

    const saveCell = (section, colId, data) => {
        const cells = { ...currentSchedule.cells };
        const originalMasterId = `${section}_${colId}`;
        
        let realMasterSection = section;
        let allSectionsToMerge = [];
        
        if (data && data.isMerging && data.mergedSections && data.mergedSections.length > 0) {
            const selectedSections = [section, ...data.mergedSections];
            const indices = selectedSections.map(s => sections.indexOf(s)).filter(i => i !== -1).sort((a, b) => a - b);
            if (indices.length > 0) {
                const minIdx = indices[0];
                const maxIdx = indices[indices.length - 1];
                allSectionsToMerge = sections.slice(minIdx, maxIdx + 1);
                realMasterSection = allSectionsToMerge[0];
            }
        }
        
        const realMasterId = `${realMasterSection}_${colId}`;

        // Clean up stale merge references for original master
        Object.keys(cells).forEach(key => {
            if (cells[key]?.masterId === originalMasterId) {
                delete cells[key];
            }
        });

        if (!data) {
            // Delete mode
            delete cells[originalMasterId];
            updateSchedule({ cells });
            return;
        }

        // Clean up stale merge references for real master if different
        if (realMasterId !== originalMasterId) {
            Object.keys(cells).forEach(key => {
                if (cells[key]?.masterId === realMasterId) {
                    delete cells[key];
                }
            });
        }

        // Reset real master cell
        cells[realMasterId] = { 
            ...data, 
            merged: false, 
            masterId: null,
            mergedSections: allSectionsToMerge.filter(s => s !== realMasterSection)
        };
        
        // Handle multi-section merge
        if (allSectionsToMerge.length > 1) {
            allSectionsToMerge.forEach(s => {
                if (s !== realMasterSection) {
                    cells[`${s}_${colId}`] = {
                        ...data,
                        merged: true,
                        masterId: realMasterId,
                        mergedSections: []
                    };
                }
            });
        }
        updateSchedule({ cells });
        setEditingCell(null);
    };

    // Pre-save check: detect same speaker+content in same column for another section
    const handleSaveCell = (section, colId, data) => {
        if (!data.mergedSections?.length && data.speaker && data.content) {
            const conflict = sections.find(s => {
                if (s === section) return false;
                const existing = currentSchedule.cells[`${s}_${colId}`];
                return existing && !existing.merged &&
                    existing.speaker === data.speaker &&
                    existing.content === data.content;
            });
            if (conflict) {
                setMergeConflict({ section, colId, data, conflictSection: conflict });
                setEditingCell(null);
                return;
            }
        }
        saveCell(section, colId, data);
    };

    const resolveMerge = (doMerge) => {
        if (!mergeConflict) return;
        const { section, colId, data, conflictSection } = mergeConflict;
        
        if (doMerge) {
            const cells = { ...currentSchedule.cells };
            const conflictCell = cells[`${conflictSection}_${colId}`];
            const masterKey = conflictCell.masterId || `${conflictSection}_${colId}`;
            const [mSection] = masterKey.split('_');
            const masterData = cells[masterKey];
            
            saveCell(mSection, colId, {
                ...masterData,
                isMerging: true,
                mergedSections: [...(masterData.mergedSections || []), section]
            });
        } else {
            saveCell(section, colId, data);
        }
        setMergeConflict(null);
    };

    const saveGuidelines = (guidelines) => {
        updateSchedule({ guidelines });
        setIsManagingGuidelines(false);
    };

    return (
        <div style={{display:'flex',flexDirection:'column',height:'100%',background:'#fff',overflow:'hidden'}} className="animate-in fade-in duration-300">
            {/* Header Controls */}
            <div className="px-6 py-4 border-b bg-slate-50 flex justify-between items-center">
                <div className="flex items-center gap-4">
                    <Button variant="ghost" icon={ArrowLeft} onClick={onBack} />
                    <div>
                        <h3 className="font-bold">Day {dayNum} Scheduler</h3>
                        <p className="text-xs text-slate-500 uppercase font-bold tracking-tight">{date}</p>
                    </div>
                </div>
                <div className="flex gap-2">
                    <Button variant="outline" icon={Copy} onClick={onClone}>Clone Schedule</Button>
                    <Button icon={Printer} onClick={onExport}>Export PDF</Button>
                    <button onClick={handleSave} style={{
                        padding:'8px 18px',borderRadius:'8px',fontWeight:700,fontSize:'14px',
                        display:'flex',alignItems:'center',gap:'8px',cursor:'pointer',transition:'all .2s',
                        background: isSaved ? '#dcfce7' : '#16a34a',
                        color: isSaved ? '#15803d' : '#fff',
                        border: isSaved ? '1px solid #86efac' : 'none'
                    }}
                        onMouseEnter={e=>{ if(!isSaved){e.currentTarget.style.background='#15803d';}}}
                        onMouseLeave={e=>{ if(!isSaved){e.currentTarget.style.background='#16a34a';}}}
                    >
                        <Save size={15} />
                        {isSaved ? '✓ Saved' : 'Save Schedule'}
                    </button>
                </div>
            </div>

            {/* Spreadsheet Grid */}
            <div style={{flex:1,overflow:'auto',background:'rgba(241,245,249,.5)'}}>
                <table style={{borderCollapse:'collapse',tableLayout:'fixed',minWidth:'100%',background:'#fff'}}>
                    <thead style={{position:'sticky',top:0,zIndex:20,boxShadow:'0 1px 2px rgba(0,0,0,.07)'}}>
                        {/* Row 1: Session Titles / Merged Headers */}
                        <tr style={{background:'#f8fafc',borderBottom:'1px solid #e2e8f0'}}>
                            <th style={{width:'192px',border:'1px solid #e2e8f0',padding:0,position:'sticky',left:0,zIndex:30,background:'#f1f5f9',height:'60px'}}>
                                <div style={{display:'flex',flexDirection:'column',justifyContent:'center',height:'100%',padding:'0 1rem', position: 'relative'}}>
                                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                                        <span style={{fontSize:'11px',textTransform:'uppercase',letterSpacing:'.05em',fontWeight:900,color:'#475569'}}>Sessions</span>
                                        <button onClick={() => setIsAddingSection(true)} style={{padding:'4px',borderRadius:'4px',color:'#2563eb',transition:'all .15s'}}>
                                            <UserPlus size={16} />
                                        </button>
                                    </div>
                                    {selectedHeaders.size >= 2 && (
                                        <div style={{position: 'absolute', inset: 0, background: '#eff6ff', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 40, animation: 'fadeIn .2s ease'}}>
                                            <button 
                                                onClick={handleMergeHeaders}
                                                style={{background: '#2563eb', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '6px', fontSize: '11px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', transition: 'all .2s'}}
                                                onMouseEnter={e => e.currentTarget.style.background = '#1d4ed8'}
                                                onMouseLeave={e => e.currentTarget.style.background = '#2563eb'}
                                            >
                                                <Combine size={14} /> Merge ({selectedHeaders.size})
                                            </button>
                                        </div>
                                    )}
                                </div>
                            </th>
                            {columns.map((col, idx) => {
                                const merge = headerMerges.find(m => m.colIds.includes(col.id));
                                // If merged, only render for the first column in the merge
                                if (merge) {
                                    if (merge.colIds[0] !== col.id) return null;
                                    return (
                                        <th key={merge.id} colSpan={merge.colIds.length} style={{border:'1px solid #e2e8f0',background:'#fff',padding:'12px',position:'relative'}}>
                                            <div className="flex flex-col items-center justify-center text-center">
                                                <div style={{fontWeight:800,color:'#1e3a8a',fontSize:'15px', marginBottom: '4px'}}>{merge.title}</div>
                                                <div style={{display:'flex', gap:'8px', justifyContent: 'center'}}>
                                                    <button onClick={() => setMergingHeaders(merge)} className="text-slate-400 hover:text-blue-600" title="Edit Merged Title"><Edit2 size={12} /></button>
                                                    <button onClick={() => removeHeaderMerge(merge.id)} className="text-slate-400 hover:text-red-500" title="Unmerge Columns"><Split size={14}/></button>
                                                </div>
                                            </div>
                                        </th>
                                    );
                                }
                                return (
                                    <th key={col.id} className="group" style={{ minWidth: '180px', border: '1px solid #e2e8f0', background: selectedHeaders.has(col.id) ? '#eff6ff' : '#fff', position: 'relative', height: '60px', padding: 0, verticalAlign: 'middle' }}>
                                        {/* Checkbox — top-left, never overlaps title */}
                                        <div style={{ position: 'absolute', top: '8px', left: '8px', zIndex: 5 }}>
                                            <input
                                                type="checkbox"
                                                checked={selectedHeaders.has(col.id)}
                                                onChange={() => toggleHeaderSelection(col.id)}
                                                style={{ width: '14px', height: '14px', cursor: 'pointer', accentColor: '#2563eb' }}
                                            />
                                        </div>

                                        {/* Edit / Delete — top-right, visible on hover */}
                                        <div className="opacity-0 group-hover:opacity-100 transition-opacity" style={{ position: 'absolute', top: '6px', right: '6px', zIndex: 5, display: 'flex', gap: '4px', alignItems: 'center' }}>
                                            <button
                                                onClick={() => setEditingHeader(col)}
                                                style={{ color: '#3b82f6', padding: '2px', background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                                                title="Edit session"
                                            >
                                                <Pencil size={12} />
                                            </button>
                                            <button
                                                onClick={(e) => { e.stopPropagation(); if (confirm('Remove this session column?')) updateSchedule({ columns: currentSchedule.columns.filter(c => c.id !== col.id) }); }}
                                                style={{ color: '#ef4444', padding: '2px', background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                                                title="Remove session"
                                            >
                                                <Trash2 size={12} />
                                            </button>
                                        </div>

                                        {/* Session Title — centered, padded to avoid button overlap */}
                                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%', padding: '0 30px', textAlign: 'center' }}>
                                            <span style={{ fontWeight: 700, color: '#334155', fontSize: '13px', lineHeight: '1.3' }}>{col.title}</span>
                                        </div>

                                        {/* Add (+) after — bottom-right, visible on hover */}
                                        <button
                                            onClick={() => addColumnAfter(col.id)}
                                            className="opacity-0 group-hover:opacity-100 transition-opacity"
                                            style={{ position: 'absolute', bottom: '5px', right: '5px', width: '18px', height: '18px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#eff6ff', color: '#2563eb', borderRadius: '50%', border: 'none', cursor: 'pointer', zIndex: 5 }}
                                            title="Add session after this"
                                        >
                                            <Plus size={10} />
                                        </button>
                                    </th>
                                );
                            })}
                            <th rowSpan={2} style={{width:'96px',border:'1px solid #e2e8f0',background:'#f8fafc', position: 'relative'}}>
                                <button onClick={addColumn} style={{padding:'8px',borderRadius:'50%',color:'#2563eb',border:'1px solid #bfdbfe'}}><Plus size={20}/></button>
                            </th>
                        </tr>
                        {/* Row 2: Specific Times sub-row */}
                        <tr style={{background:'#fff'}}>
                            <th style={{border:'1px solid #e2e8f0',position:'sticky',left:0,zIndex:30,background:'#f8fafc',height:'40px'}}>
                                <div style={{fontSize:'10px',color:'#94a3b8',padding:'0 1rem',textTransform:'uppercase',textAlign:'center'}}>Timeline</div>
                            </th>
                            {columns.map(col => (
                                <th key={`${col.id}-time`} style={{border:'1px solid #e2e8f0',padding:'0 12px',textAlign:'center'}}>
                                    <span style={{fontSize:'10px',fontWeight:700,color:'#94a3b8'}}>{formatTime12(col.startTime)} - {formatTime12(col.endTime)}</span>
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {sections.map(section => (
                            <tr key={section} style={{height:'160px'}}>
                                <td style={{border:'1px solid #e2e8f0',background:'#f8fafc',padding:'16px',position:'sticky',left:0,zIndex:10,fontWeight:900,color:'#334155',fontSize:'18px',boxShadow:'2px 0 5px rgba(0,0,0,.05)'}}>
                                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',height:'100%'}}>
                                        <span>{section}</span>
                                        <button onClick={() => removeSection(section)} style={{padding:'6px',borderRadius:'4px',color:'#fca5a5',transition:'all .15s'}}
                                            onMouseEnter={e=>e.currentTarget.style.color='#dc2626'}
                                            onMouseLeave={e=>e.currentTarget.style.color='#fca5a5'}>
                                            <MinusCircle size={18} />
                                        </button>
                                    </div>
                                </td>
                                {columns.map(col => {
                                    const cellKey = `${section}_${col.id}`;
                                    const cell = currentSchedule.cells[cellKey];
                                    if (cell?.merged) return null;
                                    
                                    // RowSpan calculation
                                    let rowSpan = 1;
                                    if (cell && !cell.merged) {
                                        sections.forEach(s => {
                                            if (s !== section && currentSchedule.cells[`${s}_${col.id}`]?.masterId === cellKey) rowSpan++;
                                        });
                                    }

                                    // Helper for consistent ColSpan calculation
                                    const getCellColSpan = (sName, cId) => {
                                        const c = currentSchedule.cells[`${sName}_${cId}`];
                                        if (!c || c.merged) return 1;
                                        const cIdx = columns.findIndex(colObj => colObj.id === cId);
                                        let span = 1;
                                        for (let i = cIdx + 1; i < columns.length; i++) {
                                            const nCol = columns[i];
                                            const nCell = currentSchedule.cells[`${sName}_${nCol.id}`];
                                            if (nCell && !nCell.merged &&
                                                nCell.type === c.type &&
                                                nCell.content === c.content &&
                                                nCell.speaker === c.speaker &&
                                                nCell.module === c.module &&
                                                nCell.examName === c.examName &&
                                                nCell.invigilator === c.invigilator &&
                                                nCell.studentRange === c.studentRange &&
                                                nCell.studentRangeFrom === c.studentRangeFrom &&
                                                nCell.studentRangeTo === c.studentRangeTo &&
                                                nCell.isDayManager === c.isDayManager &&
                                                nCell.dayManagerText === c.dayManagerText &&
                                                JSON.stringify([...(nCell.mergedSections || [])].sort()) === JSON.stringify([...(c.mergedSections || [])].sort())
                                            ) {
                                                span++;
                                            } else {
                                                break;
                                            }
                                        }
                                        return span;
                                    };

                                    const colSpan = getCellColSpan(section, col.id);

                                    // Skip logic using the same helper function
                                    let isCoveredByPrevColSpan = false;
                                    const currentColIdx = columns.findIndex(c => c.id === col.id);
                                    for (let i = 0; i < currentColIdx; i++) {
                                        const prevCol = columns[i];
                                        const prevSpan = getCellColSpan(section, prevCol.id);
                                        if (i + prevSpan > currentColIdx) {
                                            isCoveredByPrevColSpan = true;
                                            break;
                                        }
                                    }
                                    if (isCoveredByPrevColSpan) return null;

                                    return (
                                        <td key={col.id} rowSpan={rowSpan} colSpan={colSpan} style={{border:'1px solid #e2e8f0',padding:'4px',verticalAlign:'middle',background:'#fff',position:'relative'}}>
                                            {cell ? (
                                                <div
                                                    style={{
                                                        height: rowSpan > 1 ? '100%' : 'auto',
                                                        minHeight: rowSpan === 1 ? '90px' : '100%',
                                                        border:`1.5px solid ${cell.type === 'Activity' ? '#fde68a' : '#bfdbfe'}`,
                                                        borderRadius:'10px',
                                                        padding:'10px',
                                                        background: cell.type === 'Activity' ? '#fffbeb' : '#f0f9ff',
                                                        display:'flex',
                                                        flexDirection:'column',
                                                        justifyContent:'center',
                                                        position:'relative',
                                                        transition:'all .2s',
                                                        boxShadow: selectedCells.has(cellKey) ? '0 0 0 3px #eab308' : '0 1px 3px rgba(0,0,0,0.05)',
                                                        zIndex: 1
                                                    }}
                                                    onClick={(e) => {
                                                        // Prevent editor opening if clicking the edges/checkbox? 
                                                        // Actually user wants card fit, so we'll just keep the main click
                                                        setEditingCell({ section, col, cell });
                                                    }}
                                                    onMouseEnter={e => {
                                                        if (!selectedCells.has(cellKey)) {
                                                            e.currentTarget.style.boxShadow = '0 4px 12px -2px rgba(0,0,0,0.1)';
                                                            e.currentTarget.style.borderColor = cell.type === 'Activity' ? '#fcd34d' : '#3b82f6';
                                                        }
                                                    }}
                                                    onMouseLeave={e => {
                                                        if (!selectedCells.has(cellKey)) {
                                                            e.currentTarget.style.boxShadow = '0 1px 3px rgba(0,0,0,0.05)';
                                                            e.currentTarget.style.borderColor = cell.type === 'Activity' ? '#fde68a' : '#bfdbfe';
                                                        }
                                                    }}
                                                >
                                                    {/* Top area removed checkboxes */}

                                                    {/* Actions Menu (Direct Icons) */}
                                                    <div style={{position:'absolute', top:'6px', right:'6px', zIndex:100, display:'flex', gap:'4px'}} onClick={e => e.stopPropagation()}>
                                                        <button 
                                                            onClick={(e) => { 
                                                                if(confirm("Do you want to delete the schedule?")) saveCell(section, col.id, null); 
                                                            }}
                                                            className="w-7 h-7 flex items-center justify-center hover:bg-red-50 rounded-full transition-all text-slate-400 hover:text-red-500"
                                                            title="Delete Session"
                                                        >
                                                            <Trash2 size={16} />
                                                        </button>
                                                        <button 
                                                            onClick={() => setEditingCell({ section, col, cell })}
                                                            className="w-7 h-7 flex items-center justify-center hover:bg-blue-50 rounded-full transition-all text-slate-400 hover:text-blue-500"
                                                            title="Edit Session"
                                                        >
                                                            <Pencil size={14} />
                                                        </button>
                                                    </div>

                                                    {/* Session Content Box */}
                                                    <div style={{marginTop:'auto', marginBottom:'auto', display:'flex', flexDirection:'column', gap:'4px'}}>
                                                        <div style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'4px', marginTop:'12px'}}>
                                                            <span style={{
                                                                fontSize:'9px',
                                                                fontWeight:900,
                                                                textTransform:'uppercase',
                                                                padding:'2px 8px',
                                                                borderRadius:'6px',
                                                                background: cell.type==='Activity'?'#fef3c7':(cell.type==='Exam'?'#ffe4e6':'#dbeafe'),
                                                                color: cell.type==='Activity'?'#d97706':(cell.type==='Exam'?'#e11d48':'#2563eb'),
                                                                border: `1px solid ${cell.type==='Activity'?'#fde68a':(cell.type==='Exam'?'#fecaca':'#bfdbfe')}`
                                                            }}>{cell.type}</span>
                                                            <span style={{fontSize:'10px', color:'#94a3b8', fontWeight:700}}>{cell.type === 'Exam' ? '' : (cell.module || '')}</span>
                                                        </div>

                                                        <div style={{textAlign:'center'}}>
                                                            <h4 style={{fontSize:'12px', fontWeight:800, color:'#1e293b', marginBottom:'2px', lineHeight:'1.2'}}>
                                                                {cell.type === 'Exam' ? cell.examName : cell.content}
                                                            </h4>
                                                            <p style={{fontSize:'10px', color:'#64748b', fontWeight:500}}>
                                                                {cell.type === 'Exam' ? cell.invigilator : cell.speaker}
                                                            </p>
                                                        </div>

                                                        {(cell.studentRange || cell.isDayManager) && (
                                                            <div style={{display:'flex', flexWrap:'wrap', gap:'4px', justifyContent:'center', marginTop:'4px'}}>
                                                                {cell.studentRange && (
                                                                    <span style={{fontSize:'9px', background:'#f8fafc', padding:'1px 5px', borderRadius:'4px', border:'1px solid #f1f5f9', color:'#64748b'}}>ID: {cell.studentRangeFrom}-{cell.studentRangeTo}</span>
                                                                )}
                                                                {cell.isDayManager && (
                                                                    <span style={{fontSize:'9px', background:'#f8fafc', padding:'1px 5px', borderRadius:'4px', border:'1px solid #f1f5f9', color:'#64748b'}}>Manager: {cell.dayManagerText}</span>
                                                                )}
                                                            </div>
                                                        )}
                                                    </div>
                                                </div>
                                            ) : (
                                                <div
                                                    onClick={() => setEditingCell({ section, col, cell: null })}
                                                    title="Add New Schedule"
                                                    style={{height:'100%',border:'2px dashed #e2e8f0',borderRadius:'8px',display:'flex',alignItems:'center',justifyContent:'center',color:'#cbd5e1',cursor:'pointer',transition:'all .15s'}}
                                                    onMouseEnter={e=>{e.currentTarget.style.borderColor='#bfdbfe';e.currentTarget.style.color='#60a5fa';}}
                                                    onMouseLeave={e=>{e.currentTarget.style.borderColor='#e2e8f0';e.currentTarget.style.color='#cbd5e1';}}
                                                >
                                                    <Plus size={20} />
                                                </div>
                                            )}
                                        </td>
                                    );
                                })}
                                <td style={{border:'1px solid #e2e8f0',background:'rgba(248,250,252,.2)'}} />
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            {/* Dynamic Guidelines Trigger */}
            <div className="p-4 bg-slate-50 border-t flex items-center justify-between">
                <div className="flex items-center gap-4">
                    <div className="flex flex-col">
                        <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Active Guidelines</span>
                        <div className="flex gap-2 mt-1">
                            {currentSchedule.guidelines?.length > 0 ? (
                                currentSchedule.guidelines.map(g => (
                                    <span key={g.id} className="text-[11px] font-bold bg-white border border-slate-200 px-2 py-0.5 rounded-full text-slate-600 shadow-sm flex items-center gap-1">
                                        <Info size={10} className="text-blue-500" /> {g.name}
                                    </span>
                                ))
                            ) : (
                                <span className="text-xs text-slate-400 italic">No guidelines set for this date</span>
                            )}
                        </div>
                    </div>
                </div>
                <Button variant="outline" icon={FileText} onClick={() => setIsManagingGuidelines(true)}>
                    Manage Guidelines
                </Button>
            </div>

            {/* Editor Overlays */}
            {editingCell && (
                <CellEditor 
                    section={editingCell.section} 
                    sections={sections} 
                    col={editingCell.col} 
                    cell={editingCell.cell} 
                    onClose={() => setEditingCell(null)} 
                    onSave={(data) => { saveCell(editingCell.section, editingCell.col.id, data); setEditingCell(null); }} 
                    guestSpeakers={guestSpeakers}
                    onAddGuestSpeaker={(gs) => {
                        const newGs = { ...gs, id: `GS_${Date.now()}` };
                        setGuestSpeakers(prev => [...prev, newGs]);
                        return newGs;
                    }}
                />
            )}
            {editingHeader && (
                <HeaderEditor col={editingHeader} onClose={() => setEditingHeader(null)} onSave={(data) => { updateHeader(editingHeader.id, data); setEditingHeader(null); }} />
            )}
            {isAddingSection && (
                <SectionPicker existing={sections} onClose={() => setIsAddingSection(false)} onSelect={addSection} />
            )}
            {isManagingGuidelines && (
                <GuidelineManager initialData={currentSchedule.guidelines || []} onClose={() => setIsManagingGuidelines(false)} onSave={saveGuidelines} />
            )}
            {mergingHeaders && (
                <HeaderMergeModal 
                    data={mergingHeaders} 
                    onConfirm={confirmMerge} 
                    onClose={() => setMergingHeaders(null)} 
                />
            )}
            {splittingHeader && (
                <HeaderSplitModal 
                    data={splittingHeader}
                    onConfirm={confirmSplit}
                    onClose={() => setSplittingHeader(null)}
                />
            )}
            {mergeConflict && (
                <MergeConflictModal
                    conflict={mergeConflict}
                    onYes={() => resolveMerge(true)}
                    onNo={() => resolveMerge(false)}
                    onClose={() => setMergeConflict(null)}
                />
            )}
            {splittingSession && (
                <SessionSplitModal
                    data={splittingSession}
                    onClose={() => setSplittingSession(null)}
                    onConfirm={(splitPoint) => {
                        handleSplitSession(splittingSession.col.id, splitPoint);
                        setSplittingSession(null);
                    }}
                />
            )}
        </div>
    );
}

// Session Split Modal
function SessionSplitModal({ data, onClose, onConfirm }) {
    const [splitTime, setSplitTime] = useState(data.col.endTime);
    return (
        <div style={{position:'fixed',inset:0,zIndex:90,display:'flex',alignItems:'center',justifyContent:'center',background:'rgba(15,23,42,.65)',backdropFilter:'blur(4px)',padding:'16px'}}>
            <div style={{background:'#fff',borderRadius:'16px',boxShadow:'0 25px 50px -12px rgba(0,0,0,.35)',width:'100%',maxWidth:'400px',overflow:'hidden', animation:'zoomIn .2s ease'}}>
                <div style={{padding:'20px 24px 16px',borderBottom:'1px solid #f1f5f9',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <h4 style={{fontWeight:800,fontSize:'15px',color:'#0f172a'}}>Split Session Time</h4>
                    <button onClick={onClose} style={{color:'#94a3b8'}}><X size={18}/></button>
                </div>
                <div style={{padding:'24px', spaceY:'16px'}}>
                    <p style={{fontSize:'13px', color:'#64748b', marginBottom:'16px'}}>This will split the column <strong>{data.col.title}</strong> into two. Pick the end time for the first part.</p>
                    <div>
                        <label style={{fontSize:'11px',fontWeight:900,color:'#64748b',textTransform:'uppercase',display:'block',marginBottom:'8px'}}>Split at Time</label>
                        <input 
                            type="time"
                            value={splitTime} 
                            onChange={e => setSplitTime(e.target.value)}
                            style={{width:'100%',padding:'10px 12px',border:'1px solid #cbd5e1',borderRadius:'8px',fontSize:'14px',outline:'none'}}
                        />
                    </div>
                </div>
                <div style={{padding:'16px 24px 20px',display:'flex',justifyContent:'flex-end',gap:'10px', background: '#f8fafc'}}>
                    <Button variant="outline" onClick={onClose}>Cancel</Button>
                    <Button onClick={() => onConfirm(splitTime)}>Confirm Split</Button>
                </div>
            </div>
        </div>
    );
}

// --- Specific Modals ---

// Header Split Modal
function HeaderSplitModal({ data, onConfirm, onClose }) {
    const { colIds, currentNames, defaultNames } = data;
    const [names, setNames] = useState(colIds.map((_, i) => currentNames[i] || ''));

    const handleReset = () => {
        setNames(defaultNames);
    };

    const isComplete = names.every(n => n.trim().length > 0);

    return (
        <div style={{position:'fixed',inset:0,zIndex:70,display:'flex',alignItems:'center',justifyContent:'center',background:'rgba(15,23,42,.65)',backdropFilter:'blur(4px)',padding:'16px'}}>
            <div style={{background:'#fff',borderRadius:'16px',boxShadow:'0 25px 50px -12px rgba(0,0,0,.35)',width:'100%',maxWidth:'480px',overflow:'hidden', animation:'zoomIn .2s ease'}}>
                <div style={{padding:'20px 24px 16px',borderBottom:'1px solid #f1f5f9',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <div style={{display:'flex', gap:'12px', alignItems:'center'}}>
                        <div style={{width:'36px', height:'36px', borderRadius:'10px', background:'#fef2f2', display:'flex', alignItems:'center', justifyCenter:'center', flexShrink:0}}>
                            <Split size={18} style={{color:'#ef4444', margin:'0 auto'}} />
                        </div>
                        <div>
                            <h4 style={{fontWeight:800,fontSize:'15px',color:'#0f172a'}}>Unmerge & Split Headers</h4>
                            <p style={{fontSize:'12px',color:'#64748b'}}>Set individual names for each column</p>
                        </div>
                    </div>
                    <button onClick={onClose} style={{color:'#94a3b8'}}><X size={18}/></button>
                </div>
                <div style={{padding:'24px', maxHeight:'60vh', overflowY:'auto'}} className="space-y-4">
                    {colIds.map((id, i) => (
                        <div key={id} style={{marginBottom:'16px'}}>
                            <div style={{display:'flex', justifyContent:'space-between', marginBottom:'6px'}}>
                                <label style={{fontSize:'11px',fontWeight:900,color:'#64748b',textTransform:'uppercase'}}>Column {i + 1}</label>
                                <span style={{fontSize:'10px', color:'#94a3b8', fontFamily:'monospace'}}>{id}</span>
                            </div>
                            <input 
                                value={names[i]} 
                                onChange={e => {
                                    const newNames = [...names];
                                    newNames[i] = e.target.value;
                                    setNames(newNames);
                                }}
                                style={{width:'100%',padding:'10px 12px',border:'1px solid #cbd5e1',borderRadius:'8px',fontSize:'14px',outline:'none'}}
                                onFocus={e=>e.target.style.borderColor='#3b82f6'}
                                onBlur={e=>e.target.style.borderColor='#cbd5e1'}
                                placeholder="Enter column title..."
                            />
                        </div>
                    ))}
                </div>
                <div style={{padding:'16px 24px 20px',display:'flex',justifyContent:'space-between',gap:'10px', background: '#f8fafc', borderTop:'1px solid #f1f5f9'}}>
                    <Button variant="secondary" onClick={handleReset} style={{fontSize:'12px'}}>Reset to Predefined</Button>
                    <div style={{display:'flex', gap:'10px'}}>
                        <Button variant="outline" onClick={onClose}>Cancel</Button>
                        <Button onClick={() => onConfirm(names)} disabled={!isComplete}>Confirm Split</Button>
                    </div>
                </div>
                {!isComplete && (
                    <div style={{padding:'0 24px 12px', textAlign:'center'}}>
                        <p style={{fontSize:'11px', color:'#ef4444', fontWeight:600}}>⚠ All columns must have a title to split.</p>
                    </div>
                )}
            </div>
        </div>
    );
}

// Header Merge Modal
function HeaderMergeModal({ data, onConfirm, onClose }) {
    const [title, setTitle] = useState(data.title);
    return (
        <div style={{position:'fixed',inset:0,zIndex:70,display:'flex',alignItems:'center',justifyContent:'center',background:'rgba(15,23,42,.65)',backdropFilter:'blur(4px)',padding:'16px'}}>
            <div style={{background:'#fff',borderRadius:'16px',boxShadow:'0 25px 50px -12px rgba(0,0,0,.35)',width:'100%',maxWidth:'440px',overflow:'hidden'}}>
                <div style={{padding:'20px 24px 16px',borderBottom:'1px solid #f1f5f9',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <h4 style={{fontWeight:800,fontSize:'15px',color:'#0f172a'}}>Merge Columns</h4>
                    <button onClick={onClose} style={{color:'#94a3b8'}}><X size={18}/></button>
                </div>
                <div style={{padding:'24px'}}>
                    <div style={{marginBottom:'16px'}}>
                        <label style={{fontSize:'11px',fontWeight:900,color:'#64748b',textTransform:'uppercase',display:'block',marginBottom:'8px'}}>Merged Header Title</label>
                        <input 
                            value={title} 
                            onChange={e => setTitle(e.target.value)}
                            style={{width:'100%',padding:'10px 12px',border:'1px solid #cbd5e1',borderRadius:'8px',fontSize:'14px',outline:'none', boxShadow:'inset 0 1px 2px rgba(0,0,0,0.05)'}}
                            onFocus={e=>e.target.style.borderColor='#3b82f6'}
                            onBlur={e=>e.target.style.borderColor='#cbd5e1'}
                        />
                    </div>
                    <p style={{fontSize:'12px',color:'#64748b'}}>Merging columns: <span style={{fontWeight:700, color:'#334155'}}>{data.title}</span></p>
                </div>
                <div style={{padding:'12px 24px 20px',display:'flex',justifyContent:'flex-end',gap:'10px', background: '#f8fafc'}}>
                    <Button variant="outline" onClick={onClose}>Cancel</Button>
                    <Button onClick={() => onConfirm(title)}>Save Merge</Button>
                </div>
            </div>
        </div>
    );
}

// Merge Conflict Modal
function MergeConflictModal({ conflict, onYes, onNo, onClose }) {
    const { section, conflictSection, data } = conflict;
    return (
        <div style={{position:'fixed',inset:0,zIndex:70,display:'flex',alignItems:'center',justifyContent:'center',background:'rgba(15,23,42,.65)',backdropFilter:'blur(4px)',padding:'16px'}}>
            <div style={{background:'#fff',borderRadius:'16px',boxShadow:'0 25px 50px -12px rgba(0,0,0,.35)',width:'100%',maxWidth:'440px',overflow:'hidden',animation:'zoomIn .2s ease'}}>
                {/* Header */}
                <div style={{padding:'20px 24px 16px',borderBottom:'1px solid #f1f5f9',display:'flex',justifyContent:'space-between',alignItems:'flex-start'}}>
                    <div style={{display:'flex',gap:'12px',alignItems:'flex-start'}}>
                        <div style={{width:'40px',height:'40px',borderRadius:'10px',background:'#fef3c7',display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>
                            <span style={{fontSize:'20px'}}>⚠️</span>
                        </div>
                        <div>
                            <h4 style={{fontWeight:800,fontSize:'15px',color:'#0f172a',marginBottom:'4px'}}>Duplicate Speaker Detected</h4>
                            <p style={{fontSize:'13px',color:'#64748b',lineHeight:1.5}}>
                                <strong style={{color:'#2563eb'}}>{data.speaker}</strong> is already assigned to
                                Section <strong style={{color:'#0f172a'}}>{conflictSection}</strong> for the same
                                time slot with the same content.
                            </p>
                        </div>
                    </div>
                    <button onClick={onClose}
                        style={{padding:'4px',borderRadius:'6px',color:'#94a3b8',flexShrink:0,transition:'background .15s'}}
                        onMouseEnter={e=>e.currentTarget.style.background='#f1f5f9'}
                        onMouseLeave={e=>e.currentTarget.style.background=''}>
                        <X size={18} />
                    </button>
                </div>
                {/* Body */}
                <div style={{padding:'20px 24px'}}>
                    <div style={{background:'#f8fafc',border:'1px solid #e2e8f0',borderRadius:'10px',padding:'14px',marginBottom:'20px'}}>
                        <div style={{fontSize:'11px',fontWeight:700,color:'#94a3b8',textTransform:'uppercase',letterSpacing:'.05em',marginBottom:'8px'}}>Affected Sections</div>
                        <div style={{display:'flex',gap:'8px',flexWrap:'wrap'}}>
                            {[section, conflictSection].map(s => (
                                <span key={s} style={{padding:'4px 12px',borderRadius:'99px',background:'#dbeafe',color:'#1e40af',fontSize:'13px',fontWeight:700}}>{s}</span>
                            ))}
                        </div>
                        <div style={{marginTop:'10px',fontSize:'13px',color:'#475569'}}>
                            <span style={{fontWeight:600}}>Topic:</span> {data.content}
                        </div>
                    </div>
                    <p style={{fontSize:'13px',color:'#334155',fontWeight:600,marginBottom:'4px'}}>Do you want to merge these two sections?</p>
                    <p style={{fontSize:'12px',color:'#94a3b8'}}>Choosing <em>Yes</em> will visually combine them into a single merged cell in the schedule.</p>
                </div>
                {/* Footer */}
                <div style={{padding:'12px 24px 20px',display:'flex',justifyContent:'flex-end',gap:'10px'}}>
                    <button onClick={onNo}
                        style={{padding:'9px 20px',borderRadius:'8px',border:'1px solid #e2e8f0',background:'#fff',fontWeight:600,fontSize:'14px',color:'#475569',cursor:'pointer',transition:'all .15s'}}
                        onMouseEnter={e=>{e.currentTarget.style.background='#f8fafc';e.currentTarget.style.borderColor='#cbd5e1';}}
                        onMouseLeave={e=>{e.currentTarget.style.background='#fff';e.currentTarget.style.borderColor='#e2e8f0';}}>
                        No, Keep Separate
                    </button>
                    <button onClick={onYes}
                        style={{padding:'9px 20px',borderRadius:'8px',border:'none',background:'#2563eb',color:'#fff',fontWeight:700,fontSize:'14px',cursor:'pointer',transition:'background .15s',display:'flex',alignItems:'center',gap:'6px'}}
                        onMouseEnter={e=>e.currentTarget.style.background='#1d4ed8'}
                        onMouseLeave={e=>e.currentTarget.style.background='#2563eb'}>
                        <Combine size={15} /> Yes, Merge Sections
                    </button>
                </div>
            </div>
        </div>
    );
}

function GuidelineManager({ initialData, onClose, onSave }) {
    const [items, setItems] = useState(initialData.length > 0 ? initialData : [{ id: Date.now().toString(), name: '', instruction: '' }]);

    const addRow = () => setItems([...items, { id: Date.now().toString(), name: '', instruction: '' }]);
    const removeRow = (id) => setItems(items.filter(i => i.id !== id));
    const updateItem = (id, field, value) => setItems(items.map(i => i.id === id ? { ...i, [field]: value } : i));

    return (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
            <Card className="w-full max-w-2xl shadow-2xl animate-in zoom-in duration-200 overflow-hidden">
                <div className="p-4 border-b flex justify-between items-center bg-slate-50">
                    <div>
                        <h4 className="font-bold text-slate-800">Schedule Guidelines & Instructions</h4>
                        <p className="text-xs text-slate-500">These will appear at the bottom of the printed schedule</p>
                    </div>
                    <button onClick={onClose} style={{padding:'4px',borderRadius:'4px',color:'#64748b',transition:'background .15s'}}
                        onMouseEnter={e=>e.currentTarget.style.background='#e2e8f0'}
                        onMouseLeave={e=>e.currentTarget.style.background=''}><X size={20} /></button>
                </div>

                <div className="p-6 space-y-4 max-h-[60vh] overflow-auto">
                    {items.map((item) => (
                        <div key={item.id} className="flex gap-4 items-start p-4 bg-slate-50 rounded-xl border border-slate-200">
                            <div className="flex-1 space-y-3">
                                <input
                                    placeholder="Guideline Name (e.g. Dress Code)"
                                    style={{width:'100%',fontWeight:700,fontSize:'14px',background:'transparent',borderBottom:'1px solid #cbd5e1',outline:'none',paddingBottom:'4px',transition:'border-color .15s'}}
                                    value={item.name}
                                    onChange={(e) => updateItem(item.id, 'name', e.target.value)}
                                    onFocus={e=>e.target.style.borderColor='#2563eb'}
                                    onBlur={e=>e.target.style.borderColor='#cbd5e1'}
                                />
                                <textarea
                                    placeholder="Enter specific instructions..."
                                    rows={2}
                                    style={{width:'100%',fontSize:'14px',background:'#fff',border:'1px solid #e2e8f0',borderRadius:'8px',padding:'8px',outline:'none',resize:'vertical',transition:'box-shadow .15s'}}
                                    value={item.instruction}
                                    onChange={(e) => updateItem(item.id, 'instruction', e.target.value)}
                                    onFocus={e=>e.target.style.boxShadow='0 0 0 2px #bfdbfe'}
                                    onBlur={e=>e.target.style.boxShadow=''}
                                />
                            </div>
                            <button
                                onClick={() => removeRow(item.id)}
                                style={{padding:'8px',borderRadius:'8px',border:'1px solid #e2e8f0',background:'#fff',color:'#94a3b8',transition:'color .15s',boxShadow:'0 1px 2px rgba(0,0,0,.05)'}}
                                onMouseEnter={e=>e.currentTarget.style.color='#ef4444'}
                                onMouseLeave={e=>e.currentTarget.style.color='#94a3b8'}
                            >
                                <Trash2 size={16} />
                            </button>
                        </div>
                    ))}

                    <button
                        onClick={addRow}
                        style={{width:'100%',padding:'16px',border:'2px dashed #e2e8f0',borderRadius:'12px',color:'#94a3b8',display:'flex',alignItems:'center',justifyContent:'center',gap:'8px',fontWeight:700,transition:'all .15s'}}
                        onMouseEnter={e=>{e.currentTarget.style.borderColor='#bfdbfe';e.currentTarget.style.color='#2563eb';e.currentTarget.style.background='rgba(239,246,255,.5)';}}
                        onMouseLeave={e=>{e.currentTarget.style.borderColor='#e2e8f0';e.currentTarget.style.color='#94a3b8';e.currentTarget.style.background='';}}
                    >
                        <Plus size={18} /> Add Another Guideline
                    </button>
                </div>

                <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                    <Button variant="outline" onClick={onClose}>Discard Changes</Button>
                    <Button icon={Save} onClick={() => onSave(items.filter(i => i.name || i.instruction))}>Save All Guidelines</Button>
                </div>
            </Card>
        </div>
    );
}

function SectionPicker({ existing, onClose, onSelect }) {
    const available = MASTER_SECTIONS.filter(s => !existing.includes(s));
    return (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4">
            <Card className="w-full max-w-xs shadow-2xl animate-in zoom-in duration-200">
                <div className="p-4 border-b flex justify-between items-center bg-slate-50">
                    <h4 className="font-bold">Add Section</h4>
                    <button onClick={onClose} style={{padding:'4px',borderRadius:'4px',transition:'background .15s'}}
                        onMouseEnter={e=>e.currentTarget.style.background='#e2e8f0'}
                        onMouseLeave={e=>e.currentTarget.style.background=''}><X size={18} /></button>
                </div>
                <div className="p-4 space-y-2 max-h-64 overflow-auto">
                    {available.length > 0 ? available.map(s => (
                        <button key={s} onClick={() => onSelect(s)}
                            style={{width:'100%',textAlign:'left',padding:'12px 16px',borderRadius:'8px',fontWeight:500,transition:'all .15s',border:'1px solid transparent'}}
                            onMouseEnter={e=>{e.currentTarget.style.background='#eff6ff';e.currentTarget.style.color='#1d4ed8';e.currentTarget.style.borderColor='#bfdbfe';}}
                            onMouseLeave={e=>{e.currentTarget.style.background='';e.currentTarget.style.color='';e.currentTarget.style.borderColor='transparent';}}>
                            {s}
                        </button>
                    )) : <p className="text-center py-4 text-slate-400 text-sm">All sections already added.</p>}
                </div>
            </Card>
        </div>
    );
}

function HeaderEditor({ col, onClose, onSave }) {
    const [form, setForm] = useState({ ...col });
    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4">
            <Card className="w-full max-w-sm animate-in zoom-in duration-200">
                <div className="p-4 border-b flex justify-between items-center bg-slate-50">
                    <h4 className="font-bold">Edit Column</h4>
                    <button onClick={onClose} style={{padding:'4px',borderRadius:'4px',transition:'background .15s'}}
                        onMouseEnter={e=>e.currentTarget.style.background='#e2e8f0'}
                        onMouseLeave={e=>e.currentTarget.style.background=''}><X size={18} /></button>
                </div>
                <div className="p-5 space-y-4">
                    <div>
                        <label style={{fontSize:'11px',fontWeight:700,color:'#64748b',display:'block',marginBottom:'4px',textTransform:'uppercase'}}>Title</label>
                        <input style={{width:'100%',padding:'8px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}} value={form.title} onChange={e => setForm({ ...form, title: e.target.value })}
                            onFocus={e=>e.target.style.boxShadow='0 0 0 2px #bfdbfe'}
                            onBlur={e=>e.target.style.boxShadow=''} />
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                        <div>
                            <label style={{fontSize:'11px',fontWeight:700,color:'#64748b',display:'block',marginBottom:'4px',textTransform:'uppercase'}}>Start Time</label>
                            <input type="time" style={{width:'100%',padding:'8px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}} value={form.startTime} onChange={e => setForm({ ...form, startTime: e.target.value })}
                                onFocus={e=>e.target.style.boxShadow='0 0 0 2px #bfdbfe'}
                                onBlur={e=>e.target.style.boxShadow=''} />
                        </div>
                        <div>
                            <label style={{fontSize:'11px',fontWeight:700,color:'#64748b',display:'block',marginBottom:'4px',textTransform:'uppercase'}}>End Time</label>
                            <input type="time" style={{width:'100%',padding:'8px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}} value={form.endTime} onChange={e => setForm({ ...form, endTime: e.target.value })}
                                onFocus={e=>e.target.style.boxShadow='0 0 0 2px #bfdbfe'}
                                onBlur={e=>e.target.style.boxShadow=''} />
                        </div>
                    </div>
                </div>
                <div className="p-4 bg-slate-50 border-t flex justify-end gap-2">
                    <Button variant="outline" onClick={onClose}>Cancel</Button>
                    <Button onClick={() => onSave(form)}>Update</Button>
                </div>
            </Card>
        </div>
    );
}

function CellEditor({ section, sections, col, cell, onClose, onSave, guestSpeakers, onAddGuestSpeaker }) {
    const [showGuestModal, setShowGuestModal] = useState(false);
    const [form, setForm] = useState(cell || { 
        type: 'Lesson', 
        module: '', 
        content: '', 
        speaker: '', 
        mergedSections: [], 
        studentRange: false,
        studentRangeFrom: '',
        studentRangeTo: '',
        isDayManager: false,
        dayManagerText: '',
        examName: '',
        invigilator: ''
    });
    const moduleItem = MOCK_MODULES.find(m => m.code === form.module);
    const availableContent = form.module && moduleItem ? MOCK_CONTENT[moduleItem.id] || [] : [];
    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
            <Card className="w-full max-w-lg shadow-2xl animate-in fade-in zoom-in duration-200">
                <div className="p-4 bg-slate-50 border-b flex justify-between items-center">
                    <div>
                        <h4 className="font-bold text-slate-800">{col.title}</h4>
                        <p className="text-xs text-slate-500">Editing Section: <span className="font-bold text-blue-600">{section}</span></p>
                    </div>
                    <button onClick={onClose} style={{padding:'4px',borderRadius:'4px',color:'#94a3b8',transition:'background .15s'}}
                        onMouseEnter={e=>e.currentTarget.style.background='#e2e8f0'}
                        onMouseLeave={e=>e.currentTarget.style.background=''}><X size={20} /></button>
                </div>
                <div className="p-6 space-y-5">
                    <div style={{display:'flex',background:'#f1f5f9',padding:'4px',borderRadius:'8px'}}>
                        {['Lesson', 'Activity', 'Exam'].map(t => (
                            <button key={t} onClick={() => setForm({ ...form, type: t, content: '', module: '', speaker: '', examName: '', invigilator: '' })}
                                style={{flex:1,padding:'8px',borderRadius:'6px',fontSize:'14px',fontWeight:700,transition:'all .15s',
                                    background: form.type === t ? '#fff' : 'transparent',
                                    color: form.type === t ? '#2563eb' : '#64748b',
                                    boxShadow: form.type === t ? '0 1px 3px rgba(0,0,0,.1)' : 'none'}}>{t}</button>
                        ))}
                    </div>

                    {form.type === 'Exam' && (
                        <div className="space-y-4">
                            <div>
                                <label style={{fontSize:'11px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Exam Name</label>
                                <input 
                                    style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}}
                                    value={form.examName} onChange={e => setForm({ ...form, examName: e.target.value })}
                                    placeholder="Enter exam name..."
                                />
                            </div>
                            <div>
                                <label style={{fontSize:'11px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Invigilator</label>
                                <select style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}}
                                    value={form.invigilator} onChange={e => setForm({ ...form, invigilator: e.target.value })}>
                                    <option value="">Select Invigilator</option>
                                    {MOCK_SPEAKERS.map(s => <option key={s} value={s}>{s}</option>)}
                                </select>
                            </div>
                        </div>
                    )}

                    {form.type === 'Lesson' && (
                        <div className="space-y-4">
                            <div>
                                <label style={{fontSize:'11px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Module</label>
                                <select style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}}
                                    value={form.module} onChange={e => setForm({ ...form, module: e.target.value, content: '' })}
                                >
                                    <option value="">Select Module</option>
                                    {MOCK_MODULES.map(m => <option key={m.id} value={m.code}>{m.code} - {m.name}</option>)}
                                </select>
                            </div>

                            <div>
                                <label style={{fontSize:'11px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Content Topic</label>
                                <select style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px',background: !form.module ? '#f8fafc' : '#fff'}}
                                    disabled={!form.module} value={form.content} onChange={e => setForm({ ...form, content: e.target.value })}
                                >
                                    <option value="">Select Content</option>
                                    {availableContent.map(c => <option key={c} value={c}>{c}</option>)}
                                </select>
                            </div>

                            <div className="grid grid-cols-2 gap-4">
                                <div>
                                    <label style={{fontSize:'11px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Speaker Type</label>
                                    <select style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}}
                                        value={form.speakerType} onChange={e => setForm({ ...form, speakerType: e.target.value, speaker: '' })}>
                                        <option value="Speaker">Internal Speaker</option>
                                        <option value="Guest Speaker">Guest Speaker</option>
                                    </select>
                                </div>
                                <div className="relative">
                                    <label style={{fontSize:'11px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>
                                        {form.speakerType === 'Guest Speaker' ? 'Select Guest User' : 'Speaker'}
                                    </label>
                                    <div className="flex gap-2">
                                        <select style={{flex:1,padding:'10px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}}
                                            value={form.speaker} onChange={e => setForm({ ...form, speaker: e.target.value })}
                                        >
                                            <option value="">Select {form.speakerType}</option>
                                            {form.speakerType === 'Speaker' 
                                                ? MOCK_SPEAKERS.map(s => <option key={s} value={s}>{s}</option>)
                                                : guestSpeakers.map(gs => <option key={gs.id} value={gs.name}>{gs.name} ({gs.designation})</option>)
                                            }
                                        </select>
                                        {form.speakerType === 'Guest Speaker' && (
                                            <button 
                                                onClick={() => setShowGuestModal(true)}
                                                className="w-10 h-10 bg-blue-50 text-blue-600 rounded-lg flex items-center justify-center hover:bg-blue-600 hover:text-white transition-all shadow-sm flex-shrink-0"
                                                title="Add New Guest Speaker"
                                            >
                                                <UserPlus size={18} />
                                            </button>
                                        )}
                                    </div>
                                </div>
                            </div>
                        </div>
                    )}

                    {form.type === 'Activity' && (
                        <div>
                            <label style={{fontSize:'11px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Activity Type</label>
                            <select style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'6px',outline:'none',fontSize:'14px'}}
                                value={form.content} onChange={e => setForm({ ...form, content: e.target.value })}
                                onFocus={e=>e.target.style.boxShadow='0 0 0 2px #bfdbfe'}
                                onBlur={e=>e.target.style.boxShadow=''}>
                                <option value="">Select Activity</option>
                                {MOCK_ACTIVITIES.map(a => <option key={a} value={a}>{a}</option>)}
                            </select>
                        </div>
                    )}

                    {/* Student Range Section */}
                    <div className="space-y-3 p-4 bg-slate-50/50 rounded-xl border border-slate-100">
                        <label style={{display:'flex',alignItems:'center',gap:'8px',cursor:'pointer'}}>
                            <input type="checkbox" checked={form.studentRange} onChange={e => setForm({...form, studentRange: e.target.checked})} />
                            <span style={{fontSize:'13px', fontWeight:700, color:'#475569'}}>Student Range</span>
                        </label>
                        {form.studentRange && (
                            <div className="flex items-center gap-3 animate-in fade-in slide-in-from-top-1 duration-200">
                                <div style={{flex:1}}>
                                    <label style={{fontSize:'10px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'2px'}}>From ID</label>
                                    <input style={{width:'100%',padding:'8px',border:'1px solid #e2e8f0',borderRadius:'6px',fontSize:'13px'}} value={form.studentRangeFrom} onChange={e => setForm({...form, studentRangeFrom: e.target.value})} placeholder="01" />
                                </div>
                                <div style={{flex:1}}>
                                    <label style={{fontSize:'10px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'2px'}}>To ID</label>
                                    <input style={{width:'100%',padding:'8px',border:'1px solid #e2e8f0',borderRadius:'6px',fontSize:'13px'}} value={form.studentRangeTo} onChange={e => setForm({...form, studentRangeTo: e.target.value})} placeholder="80" />
                                </div>
                            </div>
                        )}
                    </div>

                    {/* Day Manager Section - Hidden for Exam */}
                    {form.type !== 'Exam' && (
                        <div className="space-y-3 p-4 bg-slate-50/50 rounded-xl border border-slate-100">
                            <label style={{display:'flex',alignItems:'center',gap:'8px',cursor:'pointer'}}>
                                <input type="checkbox" checked={form.isDayManager} onChange={e => setForm({...form, isDayManager: e.target.checked})} />
                                <span style={{fontSize:'13px', fontWeight:700, color:'#475569'}}>Day Manager Section</span>
                            </label>
                            {form.isDayManager && (
                                <div className="animate-in fade-in slide-in-from-top-1 duration-200">
                                    <label style={{fontSize:'10px',fontWeight:900,color:'#94a3b8',textTransform:'uppercase',display:'block',marginBottom:'2px'}}>Manager Details</label>
                                    <input style={{width:'100%',padding:'8px',border:'1px solid #e2e8f0',borderRadius:'6px',fontSize:'13px'}} value={form.dayManagerText} onChange={e => setForm({...form, dayManagerText: e.target.value})} placeholder="Enter text..." />
                                </div>
                            )}
                        </div>
                    )}

                    <div style={{paddingTop:'16px',borderTop:'1px solid #f1f5f9'}}>
                        <label style={{display:'flex',alignItems:'center',gap:'8px',cursor:'pointer',marginBottom:'12px'}}>
                            <input 
                                type="checkbox" 
                                checked={form.isMerging} 
                                onChange={e => setForm({ ...form, isMerging: e.target.checked })}
                                style={{width:'16px',height:'16px',accentColor:'#2563eb'}}
                            />
                            <span style={{fontSize:'13px',fontWeight:700,color:'#334155'}}>Merge section</span>
                        </label>

                        {form.isMerging && (
                            <div style={{display:'grid', gridTemplateColumns:'repeat(2, 1fr)', gap:'8px'}} className="animate-in fade-in slide-in-from-top-1 duration-200">
                                {sections.filter(s => s !== section).map(s => (
                                    <label key={s} style={{display:'flex',alignItems:'center',gap:'10px',padding:'10px',background: form.mergedSections?.includes(s) ? '#eff6ff' : '#fff', border:'1px solid', borderColor: form.mergedSections?.includes(s) ? '#bfdbfe' : '#e2e8f0', borderRadius:'8px',cursor:'pointer',transition:'all .15s'}}>
                                        <input type="checkbox" style={{width:'16px',height:'16px',accentColor:'#2563eb'}} 
                                            checked={form.mergedSections?.includes(s)} 
                                            onChange={e => {
                                                const newMerged = e.target.checked 
                                                    ? [...(form.mergedSections || []), s]
                                                    : (form.mergedSections || []).filter(item => item !== s);
                                                setForm({ ...form, mergedSections: newMerged });
                                            }} 
                                        />
                                        <span style={{fontSize:'13px', fontWeight:600, color: form.mergedSections?.includes(s) ? '#1e40af' : '#64748b'}}>{s}</span>
                                    </label>
                                ))}
                            </div>
                        )}
                    </div>
                </div>
                <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                    <Button variant="outline" onClick={onClose}>Cancel</Button>
                    <Button icon={Save} onClick={() => onSave(form)}>Save Schedule</Button>
                </div>
            </Card>

            {showGuestModal && (
                <GuestSpeakerModal 
                    onClose={() => setShowGuestModal(false)}
                    onCreate={(gs) => {
                        const newGs = onAddGuestSpeaker({ ...gs, id: `GS_${Date.now()}` });
                        setForm(prev => ({ ...prev, speaker: newGs.name, speakerType: 'Guest Speaker' }));
                        setShowGuestModal(false);
                    }}
                />
            )}
        </div>
    );
}

function GuestSpeakerModal({ onClose, onCreate }) {
    const [gs, setGs] = useState({ name: '', designation: '', email: '', phone: '' });

    return (
        <div className="fixed inset-0 bg-slate-950/40 flex items-center justify-center p-4 z-[300]">
            <Card className="w-full max-w-md shadow-2xl animate-in zoom-in-95 duration-200">
                <div className="p-4 border-b flex justify-between items-center bg-slate-50/50">
                    <h3 className="font-bold">Create Guest Speaker Account</h3>
                    <button onClick={onClose} className="p-1 hover:bg-slate-200 rounded-full transition-colors">
                        <X size={18} />
                    </button>
                </div>
                <div className="p-6 space-y-4">
                    <div>
                        <label style={{fontSize:'11px',fontWeight:900,color:'#64748b',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Full Name</label>
                        <input style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'8px',fontSize:'14px'}} value={gs.name} onChange={e => setGs({...gs, name: e.target.value})} placeholder="e.g. Dr. Jane Doe" />
                    </div>
                    <div>
                        <label style={{fontSize:'11px',fontWeight:900,color:'#64748b',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Designation</label>
                        <input style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'8px',fontSize:'14px'}} value={gs.designation} onChange={e => setGs({...gs, designation: e.target.value})} placeholder="e.g. Senior Consultant" />
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                        <div>
                            <label style={{fontSize:'11px',fontWeight:900,color:'#64748b',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Email</label>
                            <input style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'8px',fontSize:'14px'}} value={gs.email} onChange={e => setGs({...gs, email: e.target.value})} placeholder="email@example.com" />
                        </div>
                        <div>
                            <label style={{fontSize:'11px',fontWeight:900,color:'#64748b',textTransform:'uppercase',display:'block',marginBottom:'4px'}}>Phone Number</label>
                            <input style={{width:'100%',padding:'10px',border:'1px solid #e2e8f0',borderRadius:'8px',fontSize:'14px'}} value={gs.phone} onChange={e => setGs({...gs, phone: e.target.value})} placeholder="017XXXXXXXX" />
                        </div>
                    </div>
                </div>
                <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                    <Button variant="outline" onClick={onClose}>Cancel</Button>
                    <Button onClick={() => onCreate(gs)} disabled={!gs.name || !gs.phone}>Create Account</Button>
                </div>
            </Card>
        </div>
    );
}

// --- Schedule Export Module ---

// Helper: format date as dd.mm.yyyy
const fmtExportDate = (d) => d ? d.split('-').reverse().join('.') : '';
const DAY_NAMES = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

function ScheduleExport({ batch, course, date, schedule, onBack }) {
    if (!schedule) return (
        <div style={{padding:'40px',textAlign:'center',color:'#94a3b8',fontWeight:700}}>
            No schedule data found for this date.<br/>
            <span style={{fontWeight:400,fontSize:'14px'}}>Build a schedule and save it first, then export.</span>
        </div>
    );

    const dateObj = new Date(date);
    const dayName = DAY_NAMES[dateObj.getDay()];
    const startObj = new Date(batch.startDate);
    const dayNum = Math.ceil(Math.abs(dateObj - startObj) / (1000*60*60*24)) + 1;

    const isBreakCol = (col) => {
        const first = schedule.sections[0];
        const firstCell = schedule.cells[`${first}_${col.id}`];
        if (!firstCell) return false;
        
        // A break column is one where ALL sections are merged to the first one
        const mastKey = firstCell.masterId || `${first}_${col.id}`;
        return schedule.sections.every(s => {
            const c = schedule.cells[`${s}_${col.id}`];
            return c && (c.masterId === mastKey || `${s}_${col.id}` === mastKey);
        });
    };

    const cellSt = {border:'1px solid #000',padding:'6px',textAlign:'center',verticalAlign:'middle',fontSize:'11px', lineHeight: '1.2'};
    const thSt   = {border:'1px solid #000',padding:'5px',background:'#f2f2f2',textAlign:'center',verticalAlign:'middle',fontSize:'11px',fontWeight:700};

    return (
        <div style={{fontFamily:'var(--font-sans)', color: '#000'}}>
            <div id="export-toolbar" className="print:hidden" style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'24px', background: '#fff', padding: '12px', borderRadius: '8px', border: '1px solid #e2e8f0'}}>
                <Button variant="outline" icon={ArrowLeft} onClick={onBack}>Back to Editor</Button>
                <div style={{display:'flex', gap: '8px'}}>
                    <Button variant="secondary" icon={Save}>Save as Draft</Button>
                    <Button icon={Printer} onClick={() => window.print()}>Export PDF</Button>
                </div>
            </div>

            <div id="print-area" style={{background:'#fff',padding:'40px 50px',maxWidth:'1100px',margin:'0 auto',border:'1px solid #eee',boxShadow:'0 10px 30px rgba(0,0,0,0.05)',fontFamily:'"Times New Roman", Times, serif'}}>
                
                {/* Gov Header */}
                <div style={{textAlign:'center',marginBottom:'15px'}}>
                    <div style={{fontSize:'14px',fontWeight:500}}>Government of the People's Republic of Bangladesh</div>
                    <div style={{fontSize:'14px',fontWeight:700}}>National Academy for Educational Management (NAEM)</div>
                    <div style={{fontSize:'11px'}}>Ministry of Education, NAEM Road, Dhanmondi, Dhaka-1205</div>
                    <div style={{fontSize:'18px',fontWeight:700,marginTop:'10px', textDecoration: 'underline'}}>{course.name}</div>
                    <div style={{fontSize:'12px', marginTop: '4px'}}>Duration: {fmtExportDate(batch.startDate)} - {fmtExportDate(batch.endDate)}</div>
                    <div style={{fontSize:'22px',fontWeight:900,marginTop:'12px', textTransform: 'uppercase', letterSpacing: '2px'}}>Schedule</div>
                </div>

                {/* Sub Header */}
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end',margin:'15px 0 5px',fontSize:'14px',fontWeight:700}}>
                    <span>Date: {fmtExportDate(date)} ({dayName})</span>
                    <span>Day: {String(dayNum).padStart(2,'0')}</span>
                </div>

                {/* Table */}
                <table style={{width:'100%',borderCollapse:'collapse',border:'1.5px solid #000',tableLayout:'fixed'}}>
                    <thead>
                        {/* Row 1: Merged Titles */}
                        <tr>
                            <th rowSpan={4} style={{...thSt, width:'100px'}}>Sections</th>
                            {schedule.columns.map((col, idx) => {
                                const merge = (schedule.headerMerges || []).find(m => m.colIds.includes(col.id));
                                if (merge) {
                                    if (merge.colIds[0] !== col.id) return null;
                                    return <th key={merge.id} colSpan={merge.colIds.length} style={thSt}>{merge.title}</th>;
                                }
                                return <th key={col.id} style={thSt}>{col.title}</th>;
                            })}
                        </tr>
                        {/* Row 2: Sub-headings (Title) if merged? No, image has Time separately */}
                        <tr>
                            {schedule.columns.map(col => (
                                <th key={`${col.id}-sub`} style={{...thSt, fontWeight:500, background:'#fff', fontSize:'10px'}}>{col.title}</th>
                            ))}
                        </tr>
                        <tr>
                            {schedule.columns.map(col => (
                                <th key={col.id} style={{...thSt, fontWeight: 500, fontSize: '10px', background: '#fff'}}>{formatTime12(col.startTime)}</th>
                            ))}
                        </tr>
                        <tr>
                            {schedule.columns.map(col => (
                                <th key={col.id} style={{...thSt, fontWeight: 500, fontSize: '10px', background: '#fff'}}>{formatTime12(col.endTime)}</th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {schedule.sections.map((section, sIdx) => (
                            <tr key={section} style={{height:'85px'}}>
                                <td style={{...cellSt, fontWeight:700, background: '#fafafa'}}>{section}</td>
                                {schedule.columns.map(col => {
                                    const cellKey = `${section}_${col.id}`;
                                    const cell = schedule.cells[cellKey];
                                    
                                    const isFullBreak = isBreakCol(col);
                                    
                                    // If it's a full column break, only render once for first section
                                    if (isFullBreak) {
                                        if (sIdx !== 0) return null;
                                        return (
                                            <td key={col.id} rowSpan={schedule.sections.length} style={{...cellSt, background: '#f8fafc', padding: '0'}}>
                                                <div style={{writingMode:'vertical-lr', transform:'rotate(180deg)', display:'inline-block', fontWeight:700, fontSize:'12px', letterSpacing: '1px'}}>
                                                    {cell?.content || 'BREAK'}
                                                </div>
                                            </td>
                                        );
                                    }

                                    if (cell?.merged) return null;
                                    
                                    // RowSpan calculation for Export
                                    let rowSpan = 1;
                                    if (cell && !cell.merged) {
                                        schedule.sections.forEach((s) => {
                                            if (s !== section && schedule.cells[`${s}_${col.id}`]?.masterId === cellKey) {
                                                rowSpan++;
                                            }
                                        });
                                    }

                                    // Helper for consistent ColSpan calculation in Export
                                    const getExportColSpan = (sName, cId) => {
                                        const c = schedule.cells[`${sName}_${cId}`];
                                        if (!c || c.merged) return 1;
                                        const cIdx = schedule.columns.findIndex(colObj => colObj.id === cId);
                                        let span = 1;
                                        for (let i = cIdx + 1; i < schedule.columns.length; i++) {
                                            const nCol = schedule.columns[i];
                                            const nCell = schedule.cells[`${sName}_${nCol.id}`];
                                            if (nCell && !nCell.merged &&
                                                nCell.type === c.type &&
                                                nCell.content === c.content &&
                                                nCell.speaker === c.speaker &&
                                                nCell.module === c.module &&
                                                nCell.examName === c.examName &&
                                                nCell.invigilator === c.invigilator &&
                                                nCell.studentRange === c.studentRange &&
                                                nCell.studentRangeFrom === c.studentRangeFrom &&
                                                nCell.studentRangeTo === c.studentRangeTo &&
                                                nCell.isDayManager === c.isDayManager &&
                                                nCell.dayManagerText === c.dayManagerText &&
                                                JSON.stringify([...(nCell.mergedSections || [])].sort()) === JSON.stringify([...(c.mergedSections || [])].sort())
                                            ) {
                                                span++;
                                            } else {
                                                break;
                                            }
                                        }
                                        return span;
                                    };

                                    const colSpan = getExportColSpan(section, col.id);

                                    // Skip cover check for Export
                                    let isCovered = false;
                                    const cColIdx = schedule.columns.findIndex(c => c.id === col.id);
                                    for (let i = 0; i < cColIdx; i++) {
                                        const pCol = schedule.columns[i];
                                        const pSpan = getExportColSpan(section, pCol.id);
                                        if (i + pSpan > cColIdx) { isCovered = true; break; }
                                    }
                                    if (isCovered) return null;

                                    return (
                                        <td key={col.id} rowSpan={rowSpan} colSpan={colSpan} style={{...cellSt, background: cell?.type === 'Exam' ? '#fffafb' : '#fff'}}>
                                            {cell ? (
                                                <div style={{display:'flex',flexDirection:'column',gap:'4px', padding:'4px'}}>
                                                    {cell.type === 'Exam' ? (
                                                        <div style={{display:'flex', flexDirection:'column', gap:'2px'}}>
                                                            <div style={{fontWeight:900, color:'#9f1239', fontSize:'11px'}}>EXAM: {cell.examName}</div>
                                                            {cell.invigilator && <div style={{fontSize:'10px', fontStyle:'italic'}}>Invigilator: {cell.invigilator}</div>}
                                                        </div>
                                                    ) : (
                                                        <>
                                                            {cell.module && <div style={{fontWeight:700, fontSize:'10px'}}>M- {cell.module}</div>}
                                                            <div style={{fontSize: rowSpan > 1 || colSpan > 1 ? '13px' : '11px', fontWeight:600}}>{cell.content}</div>
                                                            {cell.speaker && <div style={{fontSize:'10px', color: '#444', fontStyle: 'italic'}}>{cell.speaker}</div>}
                                                        </>
                                                    )}
                                                    {cell.studentRange && <div style={{fontSize:'9px', fontWeight:700, marginTop:'2px'}}>ID: {cell.studentRangeFrom}-{cell.studentRangeTo}</div>}
                                                    {cell.isDayManager && <div style={{fontSize:'9px', fontWeight:700, color:'#166534'}}>DM: {cell.dayManagerText}</div>}
                                                </div>
                                            ) : <span style={{color: '#ccc'}}>—</span>}
                                        </td>
                                    );
                                })}
                            </tr>
                        ))}
                    </tbody>
                </table>

                {/* Guidelines */}
                <div style={{marginTop:'25px',fontSize:'13px',lineHeight:'1.6'}}>
                    {schedule.guidelines?.map(g => (
                        <div key={g.id}>
                            <span style={{fontWeight: 700}}> {g.name}: </span>
                            <span style={{fontStyle: 'italic'}}> {g.instruction} </span>
                        </div>
                    ))}
                </div>

                {/* Footer Signature */}
                <div style={{marginTop:'60px',display:'flex',justifyContent:'flex-end'}}>
                    <div style={{textAlign:'center', minWidth: '280px'}}>
                        <div style={{fontSize: '32px', fontFamily: '"Brush Script MT", cursive', color: '#333', marginBottom: '-10px'}}>
                            Coordinator
                        </div>
                        <div style={{borderTop: '1px solid #000', paddingTop: '5px'}}>
                            <div style={{fontWeight:700, fontSize: '13px', textTransform:'uppercase'}}>Course Coordinator</div>
                            <div style={{fontSize: '12px'}}>{course.name}</div>
                            <div style={{fontSize: '12px'}}>NAEM, Dhaka</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
}

// --- Clone Schedule Modal ---
function CloneModal({ schedules, onClose, onClone }) {
    const savedDates = Object.entries(schedules)
        .filter(([, s]) => s.saved)
        .sort(([a], [b]) => new Date(a) - new Date(b))
        .map(([d]) => d);

    const [fromDate, setFromDate] = useState(savedDates[0] || '');
    const [toDate, setToDate] = useState('');
    const isValid = fromDate && toDate && fromDate !== toDate;

    const inputSt = {width:'100%',padding:'10px 12px',border:'1px solid #e2e8f0',borderRadius:'8px',fontSize:'14px',outline:'none',transition:'box-shadow .15s'};

    return (
        <div style={{position:'fixed',inset:0,zIndex:80,display:'flex',alignItems:'center',justifyContent:'center',background:'rgba(15,23,42,.6)',backdropFilter:'blur(4px)',padding:'16px'}}>
            <div style={{background:'#fff',borderRadius:'16px',boxShadow:'0 25px 50px -12px rgba(0,0,0,.35)',width:'100%',maxWidth:'420px',overflow:'hidden'}}>
                {/* Header */}
                <div style={{padding:'20px 24px 16px',borderBottom:'1px solid #f1f5f9',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <div style={{display:'flex',gap:'12px',alignItems:'center'}}>
                        <div style={{width:'38px',height:'38px',borderRadius:'10px',background:'#dbeafe',display:'flex',alignItems:'center',justifyContent:'center'}}>
                            <Copy size={18} style={{color:'#2563eb'}} />
                        </div>
                        <div>
                            <h4 style={{fontWeight:800,fontSize:'15px',color:'#0f172a'}}>Clone Schedule</h4>
                            <p style={{fontSize:'12px',color:'#64748b'}}>Copy a schedule from one date to another</p>
                        </div>
                    </div>
                    <button onClick={onClose} style={{padding:'4px',borderRadius:'6px',color:'#94a3b8',transition:'background .15s'}}
                        onMouseEnter={e=>e.currentTarget.style.background='#f1f5f9'}
                        onMouseLeave={e=>e.currentTarget.style.background=''}>
                        <X size={18} />
                    </button>
                </div>
                {/* Body */}
                <div style={{padding:'20px 24px',display:'flex',flexDirection:'column',gap:'16px'}}>
                    {savedDates.length === 0 ? (
                        <div style={{padding:'24px',textAlign:'center',color:'#94a3b8',fontSize:'14px',background:'#f8fafc',borderRadius:'10px',border:'1px solid #e2e8f0'}}>
                            No saved schedules found to clone from.
                        </div>
                    ) : (
                        <>
                            <div>
                                <label style={{fontSize:'11px',fontWeight:700,color:'#64748b',textTransform:'uppercase',letterSpacing:'.05em',display:'block',marginBottom:'6px'}}>Clone From (saved schedule)</label>
                                <select value={fromDate} onChange={e => setFromDate(e.target.value)} style={inputSt}
                                    onFocus={e=>e.target.style.boxShadow='0 0 0 3px #bfdbfe'}
                                    onBlur={e=>e.target.style.boxShadow=''}>
                                    <option value="">Select date to clone from…</option>
                                    {savedDates.map(d => (
                                        <option key={d} value={d}>{fmtExportDate(d)} ({DAY_NAMES[new Date(d).getDay()]})</option>
                                    ))}
                                </select>
                            </div>
                            <div style={{display:'flex',alignItems:'center',gap:'10px',color:'#94a3b8'}}>
                                <div style={{flex:1,height:'1px',background:'#e2e8f0'}} />
                                <ArrowLeft size={14} style={{transform:'rotate(90deg)'}} />
                                <div style={{flex:1,height:'1px',background:'#e2e8f0'}} />
                            </div>
                            <div>
                                <label style={{fontSize:'11px',fontWeight:700,color:'#64748b',textTransform:'uppercase',letterSpacing:'.05em',display:'block',marginBottom:'6px'}}>Clone To (new date)</label>
                                <input type="date" value={toDate} onChange={e => setToDate(e.target.value)} style={inputSt}
                                    onFocus={e=>e.target.style.boxShadow='0 0 0 3px #bfdbfe'}
                                    onBlur={e=>e.target.style.boxShadow=''} />
                            </div>
                            {toDate && fromDate === toDate && (
                                <p style={{fontSize:'12px',color:'#ef4444',margin:0}}>⚠ Source and target date cannot be the same.</p>
                            )}
                        </>
                    )}
                </div>
                {/* Footer */}
                <div style={{padding:'12px 24px 20px',display:'flex',justifyContent:'flex-end',gap:'10px'}}>
                    <button onClick={onClose} style={{padding:'9px 20px',borderRadius:'8px',border:'1px solid #e2e8f0',background:'#fff',fontWeight:600,fontSize:'14px',color:'#475569',cursor:'pointer'}}>
                        Cancel
                    </button>
                    <button onClick={() => onClone(fromDate, toDate)} disabled={!isValid}
                        style={{padding:'9px 20px',borderRadius:'8px',border:'none',background: isValid ? '#2563eb' : '#93c5fd',color:'#fff',fontWeight:700,fontSize:'14px',cursor: isValid ? 'pointer' : 'not-allowed',display:'flex',alignItems:'center',gap:'8px',transition:'background .15s'}}
                        onMouseEnter={e=>{ if(isValid) e.currentTarget.style.background='#1d4ed8'; }}
                        onMouseLeave={e=>{ if(isValid) e.currentTarget.style.background='#2563eb'; }}>
                        <Copy size={14} /> Clone & Open Builder
                    </button>
                </div>
            </div>
        </div>
    );
}
