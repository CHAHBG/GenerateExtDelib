import { useState } from 'react'
import type { ChangeEvent } from 'react'
import { GeneratorEngineJS } from './utils/engine';

function App() {
    const [logs, setLogs] = useState<string[]>([]);
    const [loading, setLoading] = useState(false);
    const [progress, setProgress] = useState(0);
    const [files, setFiles] = useState<{ [key: string]: File | null }>({
        indiv: null, coll: null, coord_pi: null, coord_pc: null, tpl_indiv: null, tpl_coll: null
    });

    const handleFileChange = (key: string, e: ChangeEvent<HTMLInputElement>) => {
        if (e.target.files && e.target.files[0]) {
            setFiles(prev => ({ ...prev, [key]: e.target.files![0] }));
        }
    };

    const runGeneration = async () => {
        setLoading(true);
        setLogs([]);
        setProgress(0);
        const logFn = (msg: string) => {
            setLogs(prev => [...prev, msg]);
            setProgress(p => Math.min(p + 8, 95));
        };
        try {
            const engine = new GeneratorEngineJS(logFn);
            await engine.run(files as any);
            setProgress(100);
        } catch (err) {
            console.error(err);
            logFn(`❌ ERREUR: ${err}`);
        } finally {
            setLoading(false);
        }
    };

    const filesReady = Object.values(files).filter(Boolean).length >= 4;
    const fileCount = Object.values(files).filter(Boolean).length;

    return (
        <div className="min-h-screen bg-slate-50 font-sans text-slate-800">

            {/* Header */}
            <header className="bg-white border-b border-slate-200 sticky top-0 z-50">
                <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
                    <div className="flex items-center gap-4">
                        <img src="/logo.png" alt="Logo" className="w-10 h-10 object-contain" />
                        <div>
                            <h1 className="text-lg font-bold text-slate-900 tracking-tight">Génération d'extraits de délibérations</h1>
                            <p className="text-xs text-slate-500">BETPLUSAUDETAG</p>
                        </div>
                    </div>
                    <div className="flex items-center gap-3">
                        <span className={`px-3 py-1 rounded-full text-xs font-medium ${filesReady ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'}`}>
                            {filesReady ? '✓ Prêt' : `${fileCount}/4 fichiers`}
                        </span>
                    </div>
                </div>
            </header>

            <main className="max-w-6xl mx-auto px-6 py-10">
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">

                    {/* Left Column: Instructions */}
                    <div className="lg:col-span-1 space-y-6">
                        <div className="bg-white rounded-xl border border-slate-200 p-6">
                            <h2 className="font-bold text-slate-800 mb-4">Format des Fichiers Excel</h2>

                            <div className="space-y-4 text-sm">
                                <div>
                                    <h3 className="font-semibold text-slate-700 mb-2">Délibérations Individuelles</h3>
                                    <div className="bg-slate-50 rounded-lg p-3 text-xs text-slate-600 space-y-1">
                                        <div><code className="text-blue-600">nicad</code> – Clef unique</div>
                                        <div><code className="text-blue-600">Nom</code>, <code className="text-blue-600">Prenom</code></div>
                                        <div><code className="text-blue-600">Village</code>, <code className="text-blue-600">superficie</code></div>
                                        <div><code className="text-blue-600">type_usag</code></div>
                                        <div><code className="text-blue-600">Num_piece</code>, <code className="text-blue-600">Type_piece</code></div>
                                        <div><code className="text-blue-600">Date_naissance</code>, <code className="text-blue-600">Telephone</code></div>
                                    </div>
                                </div>

                                <div>
                                    <h3 className="font-semibold text-slate-700 mb-2">Délibérations Collectives</h3>
                                    <div className="bg-slate-50 rounded-lg p-3 text-xs text-slate-600 space-y-1">
                                        <div><code className="text-blue-600">nicad</code> – Clef unique</div>
                                        <div><code className="text-blue-600">Village</code>, <code className="text-blue-600">superficie</code></div>
                                        <div><code className="text-blue-600">type_usa</code></div>
                                        <div><code className="text-blue-600">Nom</code>, <code className="text-blue-600">Prenom</code>, <code className="text-blue-600">Num_piece</code></div>
                                        <div className="text-slate-400 italic">(Séparés par sauts de ligne pour plusieurs bénéficiaires)</div>
                                    </div>
                                </div>

                                <div>
                                    <h3 className="font-semibold text-slate-700 mb-2">Coordonnées (PI / PC)</h3>
                                    <div className="bg-slate-50 rounded-lg p-3 text-xs text-slate-600">
                                        <div><code className="text-blue-600">nicad</code>, <code className="text-blue-600">X</code>, <code className="text-blue-600">Y</code></div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div className="bg-blue-50 rounded-xl border border-blue-100 p-6">
                            <h3 className="font-bold text-blue-900 mb-2">📄 Tableau Double Colonne</h3>
                            <p className="text-sm text-blue-800 mb-3">Pour tenir sur une page, utilisez <code className="bg-blue-100 px-1 rounded">coords_split</code> :</p>
                            <div className="bg-white rounded-lg border border-blue-200 overflow-hidden text-xs">
                                <div className="grid grid-cols-6 bg-blue-100 font-semibold text-center p-2 text-blue-900">
                                    <div>PT</div><div>X</div><div>Y</div><div>PT</div><div>X</div><div>Y</div>
                                </div>
                                <div className="grid grid-cols-6 text-center p-2 text-slate-600">
                                    <div>P1</div><div>795127</div><div>1367770</div><div>P7</div><div>795041</div><div>1367773</div>
                                </div>
                            </div>
                        </div>
                    </div>

                    {/* Right Column: Upload & Generate */}
                    <div className="lg:col-span-2 space-y-6">

                        {/* Upload Cards */}
                        <div className="bg-white rounded-xl border border-slate-200 p-6">
                            <h2 className="font-bold text-slate-800 mb-4">1. Charger les Fichiers</h2>

                            <div className="grid grid-cols-2 gap-4 mb-6">
                                <UploadCard label="Délibérations Individuelles" file={files.indiv} onChange={(e) => handleFileChange('indiv', e)} />
                                <UploadCard label="Délibérations Collectives" file={files.coll} onChange={(e) => handleFileChange('coll', e)} />
                                <UploadCard label="Coordonnées PI" file={files.coord_pi} onChange={(e) => handleFileChange('coord_pi', e)} icon="📍" />
                                <UploadCard label="Coordonnées PC" file={files.coord_pc} onChange={(e) => handleFileChange('coord_pc', e)} icon="📍" />
                            </div>

                            <h3 className="font-semibold text-slate-700 mb-3">Modèles Word (.docx)</h3>
                            <div className="grid grid-cols-2 gap-4">
                                <UploadCard label="Modèle Individuel" file={files.tpl_indiv} onChange={(e) => handleFileChange('tpl_indiv', e)} accept=".docx" icon="📄" />
                                <UploadCard label="Modèle Collectif" file={files.tpl_coll} onChange={(e) => handleFileChange('tpl_coll', e)} accept=".docx" icon="📄" />
                            </div>
                        </div>

                        {/* Generate Section */}
                        <div className="bg-white rounded-xl border border-slate-200 p-6">
                            <h2 className="font-bold text-slate-800 mb-4">2. Générer les Extraits</h2>

                            <button
                                onClick={runGeneration}
                                disabled={!filesReady || loading}
                                className={`w-full py-4 rounded-lg font-semibold text-base transition-all duration-200
                  ${!filesReady || loading
                                        ? 'bg-slate-100 text-slate-400 cursor-not-allowed'
                                        : 'bg-slate-900 text-white hover:bg-slate-800 active:scale-[0.99]'
                                    }`}
                            >
                                {loading ? (
                                    <span className="flex items-center justify-center gap-3">
                                        <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" /><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" /></svg>
                                        Génération en cours...
                                    </span>
                                ) : "Lancer la Génération"}
                            </button>

                            {/* Progress Bar */}
                            {loading && (
                                <div className="mt-4">
                                    <div className="h-2 bg-slate-100 rounded-full overflow-hidden">
                                        <div className="h-full bg-blue-600 transition-all duration-300" style={{ width: `${progress}%` }}></div>
                                    </div>
                                    <p className="text-xs text-slate-500 mt-1 text-center">{progress}%</p>
                                </div>
                            )}

                            {/* Logs */}
                            {logs.length > 0 && (
                                <div className="mt-4 bg-slate-50 rounded-lg p-4 max-h-60 overflow-y-auto border border-slate-100">
                                    <div className="font-mono text-xs space-y-1">
                                        {logs.map((log, i) => (
                                            <div key={i} className={`${log.includes('❌') ? 'text-red-600' : log.includes('✅') ? 'text-emerald-600 font-semibold' : log.includes('✓') ? 'text-emerald-600' : 'text-slate-600'}`}>
                                                {log}
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            )}

                            {progress === 100 && (
                                <div className="mt-4 p-4 bg-emerald-50 rounded-lg border border-emerald-200 text-center">
                                    <p className="text-emerald-800 font-semibold">✅ Génération terminée !</p>
                                    <p className="text-emerald-600 text-sm mt-1">Le ZIP contient tous les extraits individuels + un fichier fusionné.</p>
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            </main>

            {/* Footer */}
            <footer className="border-t border-slate-200 bg-white mt-16 py-6 text-center text-slate-400 text-sm">
                © 2025 BETPLUSAUDETAG - Design by CABG
            </footer>
        </div>
    )
}

interface UploadCardProps {
    label: string;
    file: File | null;
    onChange: (e: ChangeEvent<HTMLInputElement>) => void;
    accept?: string;
    icon?: string;
}

function UploadCard({ label, file, onChange, accept = ".xlsx,.xls", icon = "📎" }: UploadCardProps) {
    return (
        <label className="block cursor-pointer">
            <div className={`p-4 rounded-lg border-2 border-dashed transition-all duration-200 ${file
                ? 'border-emerald-300 bg-emerald-50'
                : 'border-slate-200 hover:border-blue-300 hover:bg-blue-50/50'
                }`}>
                <div className="flex items-center gap-3">
                    <span className={`text-xl ${file ? 'text-emerald-500' : 'text-slate-400'}`}>{icon}</span>
                    <div className="flex-1 min-w-0">
                        <p className={`text-sm font-medium truncate ${file ? 'text-emerald-800' : 'text-slate-600'}`}>
                            {file ? file.name : label}
                        </p>
                        {!file && <p className="text-xs text-slate-400">Cliquez pour sélectionner</p>}
                    </div>
                    {file && <span className="text-emerald-500 text-sm">✓</span>}
                </div>
            </div>
            <input type="file" onChange={onChange} accept={accept} className="hidden" />
        </label>
    )
}

export default App
