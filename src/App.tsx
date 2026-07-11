/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useRef, ChangeEvent } from "react";
import * as XLSX from "xlsx";
import {
  Play,
  Pause,
  RotateCcw,
  RefreshCw,
  FileSpreadsheet,
  Upload,
  Download,
  Flame,
  Volume2,
  Calendar,
  ChevronRight,
  Share2,
  CheckCircle2,
  Sparkles,
  Info,
  Smartphone
} from "lucide-react";

// Types for the workout application
interface WorkoutItem {
  dia: number;
  ejercicio: string;
  tiempo: number; // in seconds
  tiempoOriginal: string; // original Excel representation
}

// Default preloaded 7-day routine to make the app work out-of-the-box
const DEFAULT_ROUTINE: WorkoutItem[] = [
  // Día 1: Piernas y Core
  { dia: 1, ejercicio: "Calentamiento: Trote Ligero", tiempo: 60, tiempoOriginal: "01:00" },
  { dia: 1, ejercicio: "Sentadillas Goblet", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 1, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 1, ejercicio: "Zancadas Alternas (Lunges)", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 1, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 1, ejercicio: "Puente de Glúteos", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 1, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 1, ejercicio: "Plancha Abdominal", tiempo: 60, tiempoOriginal: "01:00" },
  { dia: 1, ejercicio: "Estiramientos Flexibilidad", tiempo: 90, tiempoOriginal: "01:30" },

  // Día 2: Empuje (Pecho, Hombro, Tríceps)
  { dia: 2, ejercicio: "Calentamiento Articular", tiempo: 60, tiempoOriginal: "01:00" },
  { dia: 2, ejercicio: "Flexiones de Pecho (Push-ups)", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 2, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 2, ejercicio: "Press Militar con Mancuernas", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 2, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 2, ejercicio: "Fondos en Silla (Dips)", tiempo: 40, tiempoOriginal: "00:40" },
  { dia: 2, ejercicio: "Descanso", tiempo: 20, tiempoOriginal: "00:20" },
  { dia: 2, ejercicio: "Flexiones de Pecho Diamante", tiempo: 30, tiempoOriginal: "00:30" },
  { dia: 2, ejercicio: "Enfriamiento y Elongación", tiempo: 60, tiempoOriginal: "01:00" },

  // Día 3: Cardio HIIT
  { dia: 3, ejercicio: "Calentamiento: Jumping Jacks", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 3, ejercicio: "Burpees de Alta Intensidad", tiempo: 30, tiempoOriginal: "00:30" },
  { dia: 3, ejercicio: "Descanso Activo (Caminar)", tiempo: 30, tiempoOriginal: "00:30" },
  { dia: 3, ejercicio: "Mountain Climbers", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 3, ejercicio: "Descanso Activo (Caminar)", tiempo: 30, tiempoOriginal: "00:30" },
  { dia: 3, ejercicio: "Skips de Rodillas Altas", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 3, ejercicio: "Descanso Activo (Caminar)", tiempo: 30, tiempoOriginal: "00:30" },
  { dia: 3, ejercicio: "Abdominales Bicicleta", tiempo: 60, tiempoOriginal: "01:00" },

  // Día 4: Tracción (Espalda, Bíceps)
  { dia: 4, ejercicio: "Calentamiento: Rotaciones", tiempo: 60, tiempoOriginal: "01:00" },
  { dia: 4, ejercicio: "Dominadas o Jalón con Banda", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 4, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 4, ejercicio: "Curl de Bíceps con Mancuernas", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 4, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 4, ejercicio: "Remo al Pecho con Banda", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 4, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 4, ejercicio: "Plancha Spiderman", tiempo: 45, tiempoOriginal: "00:45" },

  // Día 5: Piernas Plio e Isométricos
  { dia: 5, ejercicio: "Calentamiento Dinámico", tiempo: 60, tiempoOriginal: "01:00" },
  { dia: 5, ejercicio: "Sentadillas con Salto (Plio)", tiempo: 30, tiempoOriginal: "00:30" },
  { dia: 5, ejercicio: "Descanso", tiempo: 20, tiempoOriginal: "00:20" },
  { dia: 5, ejercicio: "Sentadilla Isométrica (Wall Sit)", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 5, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 5, ejercicio: "Zancadas con Salto Alternadas", tiempo: 30, tiempoOriginal: "00:30" },
  { dia: 5, ejercicio: "Plancha Abdominal Lateral", tiempo: 60, tiempoOriginal: "01:00" },

  // Día 6: Core Explosivo
  { dia: 6, ejercicio: "Calentamiento General", tiempo: 60, tiempoOriginal: "01:00" },
  { dia: 6, ejercicio: "Abdominales de Mariposa", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 6, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 6, ejercicio: "Plancha con Rotación de Cadera", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 6, ejercicio: "Descanso", tiempo: 15, tiempoOriginal: "00:15" },
  { dia: 6, ejercicio: "Toques de Tobillo Cruzado", tiempo: 45, tiempoOriginal: "00:45" },
  { dia: 6, ejercicio: "Superman Isométrico", tiempo: 60, tiempoOriginal: "01:00" },

  // Día 7: Recuperación Activa
  { dia: 7, ejercicio: "Trote Muy Ligero o Caminata", tiempo: 300, tiempoOriginal: "05:00" },
  { dia: 7, ejercicio: "Estiramiento Tren Superior", tiempo: 120, tiempoOriginal: "02:00" },
  { dia: 7, ejercicio: "Estiramiento Tren Inferior", tiempo: 120, tiempoOriginal: "02:00" },
  { dia: 7, ejercicio: "Relajación y Respiraciones", tiempo: 180, tiempoOriginal: "03:00" }
];

export default function App() {
  const getDayName = (dia: number): string => {
    const dayNames: Record<number, string> = {
      1: "Lunes",
      2: "Martes",
      3: "Miércoles",
      4: "Jueves",
      5: "Viernes",
      6: "Sábado",
      7: "Domingo"
    };
    return dayNames[dia] || `Día ${dia}`;
  };

  // Load initial routine from localStorage or use default
  const [routine, setRoutine] = useState<WorkoutItem[]>(() => {
    const saved = localStorage.getItem("personal_workout_routine");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error("No se pudo cargar la rutina desde localStorage", e);
      }
    }
    return DEFAULT_ROUTINE;
  });

  // Extract unique available days based on loaded routine
  const availableDays: number[] = Array.from(new Set(routine.map((item) => item.dia))).sort((a: number, b: number) => a - b) as number[];

  // Active workout states
  const [selectedDay, setSelectedDay] = useState<number>(() => {
    return availableDays.length > 0 ? availableDays[0] : 1;
  });

  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [timeLeft, setTimeLeft] = useState<number>(0);
  const [isPlaying, setIsPlaying] = useState<boolean>(false);
  const [excelError, setExcelError] = useState<string | null>(null);
  const [excelSuccess, setExcelSuccess] = useState<boolean>(false);

  // High-precision references for Date.now() subtraction
  const startTimeRef = useRef<number>(0);
  const timeLeftOnPauseRef = useRef<number>(0);
  const timerIntervalRef = useRef<NodeJS.Timeout | null>(null);

  // Filter exercises for selected day
  const activeExercises = routine.filter((item) => item.dia === selectedDay);
  const currentExercise = activeExercises[currentIndex];

  // Initialize time left when active exercise or index changes
  useEffect(() => {
    if (currentExercise) {
      setTimeLeft(currentExercise.tiempo);
      timeLeftOnPauseRef.current = currentExercise.tiempo;
    } else {
      setTimeLeft(0);
      timeLeftOnPauseRef.current = 0;
    }
    // Pause the timer on index or day change to prevent sudden countdowns
    setIsPlaying(false);
  }, [currentIndex, selectedDay, routine]);

  // Handle timer playing/pausing state
  useEffect(() => {
    if (isPlaying) {
      // Setup high precision timestamp delta subtraction
      startTimeRef.current = Date.now();
      timeLeftOnPauseRef.current = timeLeft;

      timerIntervalRef.current = setInterval(() => {
        const elapsedMs = Date.now() - startTimeRef.current;
        const elapsedSeconds = Math.floor(elapsedMs / 1000);
        const nextTimeLeft = Math.max(0, timeLeftOnPauseRef.current - elapsedSeconds);

        if (nextTimeLeft !== timeLeft) {
          setTimeLeft(nextTimeLeft);
        }

        if (nextTimeLeft === 0) {
          // Timer reached 00:00!
          clearInterval(timerIntervalRef.current!);
          timerIntervalRef.current = null;
          playAlertSound();
          handleExerciseEnd();
        }
      }, 100); // Check every 100ms for precision
    } else {
      if (timerIntervalRef.current) {
        clearInterval(timerIntervalRef.current);
        timerIntervalRef.current = null;
      }
    }

    return () => {
      if (timerIntervalRef.current) {
        clearInterval(timerIntervalRef.current);
      }
    };
  }, [isPlaying, currentIndex, selectedDay]);

  // Automatically unlocks Web Audio API on first tap
  useEffect(() => {
    const unlockAudio = () => {
      try {
        const AudioContextClass = window.AudioContext || (window as any).webkitAudioContext;
        if (!AudioContextClass) return;
        const ctx = new AudioContextClass();
        if (ctx.state === "suspended") {
          ctx.resume().then(() => {
            // Play quick inaudible tone to unlock
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.frequency.setValueAtTime(1000, ctx.currentTime);
            gain.gain.setValueAtTime(0.001, ctx.currentTime);
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.start();
            osc.stop(ctx.currentTime + 0.05);
          });
        }
      } catch (e) {
        console.warn("AudioContext unlock failed", e);
      }
      // Remove listener immediately after first gesture
      window.removeEventListener("click", unlockAudio);
      window.removeEventListener("touchstart", unlockAudio);
    };

    window.addEventListener("click", unlockAudio);
    window.addEventListener("touchstart", unlockAudio);

    return () => {
      window.removeEventListener("click", unlockAudio);
      window.removeEventListener("touchstart", unlockAudio);
    };
  }, []);

  // Play native beep sound using JavaScript AudioContext API (440Hz, 0.5s)
  const playAlertSound = () => {
    try {
      const AudioContextClass = window.AudioContext || (window as any).webkitAudioContext;
      if (!AudioContextClass) return;
      const ctx = new AudioContextClass();
      
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      
      osc.type = "sine";
      osc.frequency.setValueAtTime(440, ctx.currentTime); // Standard 440Hz A note
      
      // Volume envelope: smooth fade out to avoid speaker pops
      gain.gain.setValueAtTime(0.4, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.5);
      
      osc.connect(gain);
      gain.connect(ctx.destination);
      
      osc.start();
      osc.stop(ctx.currentTime + 0.5);
    } catch (e) {
      console.warn("AudioContext beep failed. Browser policies might require user interaction first.", e);
    }
  };

  // Logic to handle countdown expiration
  const handleExerciseEnd = () => {
    if (currentIndex < activeExercises.length - 1) {
      // Go to next exercise immediately and auto-start
      setCurrentIndex((prev) => prev + 1);
      // Wait for React state update, and trigger start
      setTimeout(() => {
        setIsPlaying(true);
      }, 100);
    } else {
      // Last exercise completed
      setIsPlaying(false);
      setCurrentIndex(0);
      alert("¡Enhorabuena! Has completado el entrenamiento de hoy. 💪🎉");
    }
  };

  // Controller Handlers
  const handleTogglePlay = () => {
    if (!currentExercise) return;
    setIsPlaying(!isPlaying);
  };

  const handleRepeatExercise = () => {
    if (!currentExercise) return;
    // Reset timer back to active exercise's original time
    setIsPlaying(false);
    setTimeout(() => {
      setTimeLeft(currentExercise.tiempo);
      timeLeftOnPauseRef.current = currentExercise.tiempo;
    }, 50);
  };

  const handleResetDay = () => {
    setIsPlaying(false);
    setCurrentIndex(0);
    if (currentExercise) {
      setTimeLeft(currentExercise.tiempo);
      timeLeftOnPauseRef.current = currentExercise.tiempo;
    }
  };

  const handleSelectDay = (dayNum: number) => {
    setSelectedDay(dayNum);
    setCurrentIndex(0);
  };

  const handleSelectExerciseDirectly = (index: number) => {
    setIsPlaying(false);
    setCurrentIndex(index);
  };

  // Excel File Parsing & Validation
  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setExcelError(null);
    setExcelSuccess(false);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = event.target?.result;
        if (!data) throw new Error("No se obtuvieron datos del archivo.");

        const workbook = XLSX.read(data, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json<any>(worksheet);

        if (jsonData.length === 0) {
          throw new Error("El archivo Excel no contiene ninguna fila de datos.");
        }

        const parsedItems: WorkoutItem[] = [];

        for (let i = 0; i < jsonData.length; i++) {
          const row = jsonData[i];

          // Normalize keys (remove tildes, spacing, lowercase)
          const normalizeKey = (str: string) => 
            str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

          const dayKey = Object.keys(row).find((k) => normalizeKey(k) === "dia" || normalizeKey(k) === "day");
          const nameKey = Object.keys(row).find(
            (k) => normalizeKey(k) === "ejercicio" || normalizeKey(k) === "exercise" || normalizeKey(k) === "name"
          );
          const timeKey = Object.keys(row).find(
            (k) => normalizeKey(k) === "tiempo" || normalizeKey(k) === "time" || normalizeKey(k) === "duration"
          );

          if (!dayKey || !nameKey || !timeKey) {
            throw new Error(
              `Fila ${i + 2}: El archivo Excel debe contener exactamente las columnas 'Día', 'Ejercicio' y 'Tiempo'. Asegúrate de que los encabezados estén escritos en la primera fila.`
            );
          }

          const rawDay = parseInt(row[dayKey]);
          if (isNaN(rawDay) || rawDay < 1 || rawDay > 7) {
            continue; // Skip lines with invalid day range
          }

          const exerciseName = String(row[nameKey] || "").trim();
          if (!exerciseName) {
            continue; // Skip lines with empty names
          }

          const rawTime = row[timeKey];
          let parsedSeconds = 0;
          let timeDisplayStr = "";

          if (typeof rawTime === "number") {
            // Excel decimal time fraction (e.g. 0.001157 is approx 1m40s)
            if (rawTime > 0 && rawTime < 1) {
              parsedSeconds = Math.round(rawTime * 86400);
              const m = Math.floor(parsedSeconds / 60);
              const s = parsedSeconds % 60;
              timeDisplayStr = `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
            } else {
              parsedSeconds = Math.round(rawTime);
              const m = Math.floor(parsedSeconds / 60);
              const s = parsedSeconds % 60;
              timeDisplayStr = `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
            }
          } else if (typeof rawTime === "string") {
            const timeStr = rawTime.trim();
            timeDisplayStr = timeStr;
            if (timeStr.includes(":")) {
              const parts = timeStr.split(":");
              if (parts.length === 2) {
                const mins = parseInt(parts[0], 10);
                const secs = parseInt(parts[1], 10);
                if (!isNaN(mins) && !isNaN(secs)) {
                  parsedSeconds = mins * 60 + secs;
                }
              } else if (parts.length === 3) {
                const hrs = parseInt(parts[0], 10);
                const mins = parseInt(parts[1], 10);
                const secs = parseInt(parts[2], 10);
                if (!isNaN(hrs) && !isNaN(mins) && !isNaN(secs)) {
                  parsedSeconds = hrs * 3600 + mins * 60 + secs;
                }
              }
            } else {
              const secondsInt = parseInt(timeStr, 10);
              if (!isNaN(secondsInt)) {
                parsedSeconds = secondsInt;
              }
            }
          }

          if (parsedSeconds <= 0) {
            parsedSeconds = 30; // Fallback to 30 seconds
            timeDisplayStr = "00:30";
          }

          parsedItems.push({
            dia: rawDay,
            ejercicio: exerciseName,
            tiempo: parsedSeconds,
            tiempoOriginal: timeDisplayStr
          });
        }

        if (parsedItems.length === 0) {
          throw new Error("No se procesaron filas válidas. Verifica las columnas 'Día', 'Ejercicio' y 'Tiempo'.");
        }

        // Save parsed routine to State & LocalStorage
        setRoutine(parsedItems);
        localStorage.setItem("personal_workout_routine", JSON.stringify(parsedItems));
        
        // Reset navigation to the first available day and first item
        const parsedDays = Array.from(new Set(parsedItems.map((item) => item.dia))).sort((a, b) => a - b);
        setSelectedDay(parsedDays[0] || 1);
        setCurrentIndex(0);
        setExcelSuccess(true);
        setTimeout(() => setExcelSuccess(false), 5000); // Clear badge
      } catch (err: any) {
        setExcelError(err.message || "Error al leer o procesar el archivo Excel.");
      }
    };

    reader.onerror = () => {
      setExcelError("Error al abrir el archivo.");
    };
    reader.readAsBinaryString(file);
  };

  // Reset routine back to default sample
  const handleResetToDefault = () => {
    if (confirm("¿Estás seguro de que deseas restablecer la rutina de prueba de 7 días predeterminada? Esto borrará tu rutina cargada.")) {
      setRoutine(DEFAULT_ROUTINE);
      localStorage.setItem("personal_workout_routine", JSON.stringify(DEFAULT_ROUTINE));
      setSelectedDay(1);
      setCurrentIndex(0);
      setExcelSuccess(true);
      setTimeout(() => setExcelSuccess(false), 2000);
    }
  };

  // Helper: Generates a pre-populated Excel workbook and triggers direct download
  const handleDownloadSampleExcel = () => {
    const sampleData = [
      { "Día": 1, "Ejercicio": "Calentamiento: Trote Suave", "Tiempo": "01:00" },
      { "Día": 1, "Ejercicio": "Flexiones de Pecho (Push-ups)", "Tiempo": "00:45" },
      { "Día": 1, "Ejercicio": "Descanso", "Tiempo": "00:15" },
      { "Día": 1, "Ejercicio": "Sentadillas Profundas", "Tiempo": "00:45" },
      { "Día": 1, "Ejercicio": "Descanso", "Tiempo": "00:15" },
      { "Día": 1, "Ejercicio": "Plancha Abdominal de Acero", "Tiempo": "01:00" },
      { "Día": 1, "Ejercicio": "Estiramiento Estático", "Tiempo": "01:30" },

      { "Día": 2, "Ejercicio": "Jumping Jacks de Hombro", "Tiempo": 45 },
      { "Día": 2, "Ejercicio": "Fondos de Tríceps (Dips)", "Tiempo": 45 },
      { "Día": 2, "Ejercicio": "Descanso", "Tiempo": 15 },
      { "Día": 2, "Ejercicio": "Zancadas Alternas", "Tiempo": 45 },
      { "Día": 2, "Ejercicio": "Plancha Lateral Core", "Tiempo": 60 }
    ];

    const ws = XLSX.utils.json_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "RutinaEjemplo");
    XLSX.writeFile(wb, "Rutina_Entrenamiento_Ejemplo.xlsx");
  };

  // Helper: Compiles a beautiful, standalone HTML version of the app
  // It loads Tailwind, SheetJS and Font from CDN and recreates the exact UI & stopwatch logic.
  const handleExportSingleHTML = () => {
    const compiledHTML = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <title>PROTIMER - Entrenamiento Personal</title>
  
  <!-- Tailwind CSS Play CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  
  <!-- Google Fonts: Outfit, Space Grotesk & JetBrains Mono -->
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Space+Grotesk:wght@500;700&family=JetBrains+Mono:wght@700&display=swap" rel="stylesheet">
  
  <!-- SheetJS CDN -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <style>
    body {
      font-family: 'Outfit', sans-serif;
    }
    .font-display {
      font-family: 'Space Grotesk', sans-serif;
    }
    .font-mono {
      font-family: 'JetBrains Mono', monospace;
    }
    /* Custom thin visual scrollbar */
    .custom-scroll::-webkit-scrollbar {
      width: 4px;
      height: 4px;
    }
    .custom-scroll::-webkit-scrollbar-track {
      background: rgba(255, 255, 255, 0.03);
    }
    .custom-scroll::-webkit-scrollbar-thumb {
      background: #f97316;
      border-radius: 99px;
    }
  </style>
</head>
<body class="bg-[#0a0a0a] text-white min-h-screen flex flex-col antialiased selection:bg-[#f97316] selection:text-black p-4 md:p-8">
  
  <div class="w-full max-w-md mx-auto flex-1 flex flex-col justify-between">
    
    <!-- HEADER -->
    <header class="flex items-center justify-between py-4 border-b border-[#1a1a1a] mb-4">
      <div>
        <h1 class="text-xs tracking-[0.3em] font-bold text-gray-500 uppercase">PROTIMER / STANDALONE</h1>
        <p class="text-lg font-light tracking-tight text-white" id="headerDayDisplay">Día 01 <span class="text-orange-500 font-mono">• Rutina Activa</span></p>
      </div>
      
      <div class="flex items-center gap-1.5">
        <!-- Input file button -->
        <label class="flex items-center gap-1.5 px-3 py-1.5 bg-[#1a1a1a] border border-[#333] hover:border-orange-500 rounded-md text-[10px] font-mono tracking-wider uppercase text-slate-300 hover:text-white cursor-pointer transition-all">
          <svg class="w-3.5 h-3.5 text-orange-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" /></svg>
          <span>Excel</span>
          <input type="file" id="excelFile" class="hidden" accept=".xlsx, .xls">
        </label>
      </div>
    </header>

    <!-- CONTENT WRAPPER -->
    <main class="flex-1 flex flex-col gap-4">
      
      <!-- Day Pills Selector -->
      <div id="daySelectorContainer" class="flex gap-1.5 overflow-x-auto pb-1 custom-scroll">
        <!-- Injected Dynamically -->
      </div>

      <!-- TIMER DISPLAY CARD -->
      <div id="timerCard" class="relative bg-[#0d0d0d] rounded-xl border border-[#1a1a1a] p-6 shadow-xl shadow-black/40 overflow-hidden flex flex-col items-center justify-center">
        
        <!-- Active indicator -->
        <div class="absolute top-4 left-4 flex items-center gap-2">
          <div id="standaloneStatusDot" class="w-2 h-2 rounded-full bg-orange-500 animate-pulse"></div>
          <span class="text-[9px] tracking-[0.2em] font-mono uppercase text-gray-500" id="statusBadge">EJERCICIO</span>
        </div>

        <!-- Exercise Title -->
        <h2 id="exerciseTitle" class="text-2xl font-bold text-center tracking-tight mt-6 max-h-16 overflow-hidden text-white uppercase px-4">
          Carga una Rutina Excel
        </h2>
        
        <!-- MM:SS Display inside SVG Ring -->
        <div class="relative w-52 h-52 md:w-56 md:h-56 my-3 flex items-center justify-center">
          <svg class="w-full h-full -rotate-90" viewBox="0 0 100 100">
            <!-- Background circle -->
            <circle cx="50" cy="50" r="44" stroke="#1a1a1a" stroke-width="6" fill="none" />
            <!-- Active progressive circle -->
            <circle id="progressCircle" cx="50" cy="50" r="44" stroke="#f97316" stroke-width="6" fill="none" stroke-linecap="round" stroke-dasharray="276.46" stroke-dashoffset="0" class="transition-all duration-200" />
          </svg>
          
          <!-- Large digital numbers -->
          <div class="absolute flex flex-col items-center justify-center">
            <span id="countdownDisplay" class="text-5xl font-mono font-bold text-white tracking-tighter">00:00</span>
            <span id="progressPercent" class="text-[9px] text-gray-500 tracking-[0.25em] font-mono uppercase mt-1">LISTO</span>
          </div>
        </div>

        <!-- Progress slider bar -->
        <div class="w-full max-w-[220px] h-1 bg-[#1a1a1a] rounded-full overflow-hidden relative mb-4">
          <div id="horizontalProgress" class="h-full bg-orange-500 transition-all duration-300" style="width: 0%"></div>
        </div>

        <!-- Next exercise preview -->
        <div class="w-full bg-[#0a0a0a] border border-[#1a1a1a] rounded-lg py-2.5 px-3 flex items-center justify-center gap-1.5 text-center">
          <span class="text-[10px] text-gray-500 tracking-wider font-mono uppercase">SIGUIENTE:</span>
          <span id="nextExerciseDisplay" class="text-[11px] font-bold text-white truncate">Ninguno</span>
        </div>
      </div>

      <!-- TACTILE BIG TOUCH CONTROLLER PANEL -->
      <div class="grid grid-cols-3 items-center gap-3 py-2">
        <!-- REINICIAR DIA (LEFT) -->
        <button id="btnResetDay" class="flex flex-col items-center justify-center gap-1.5 p-3.5 bg-[#0d0d0d] border border-[#1a1a1a] hover:border-gray-700 hover:text-white rounded-xl active:scale-95 transition-all text-gray-400 cursor-pointer select-none">
          <svg class="w-4 h-4 text-gray-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M11 15l-3-3m0 0l3-3m-3 3h8M3 12a9 9 0 1118 0 9 9 0 01-18 0z" /></svg>
          <span class="text-[9px] font-mono tracking-widest uppercase">REINICIAR</span>
        </button>

        <!-- MAIN START / PAUSE BUTTON (CENTER - HUGE) -->
        <button id="btnPlayPause" class="h-16 w-16 mx-auto bg-[#f97316] hover:bg-[#ff8c3b] rounded-full flex items-center justify-center text-black active:scale-95 transition-all cursor-pointer border-4 border-[#0a0a0a] shadow-lg shadow-orange-500/20">
          <svg id="playIconSvg" class="w-6 h-6 fill-current ml-0.5" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
          <svg id="pauseIconSvg" class="w-6 h-6 fill-current hidden" viewBox="0 0 24 24"><path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/></svg>
        </button>

        <!-- REPETIR EJERCICIO (RIGHT) -->
        <button id="btnRepeatExercise" class="flex flex-col items-center justify-center gap-1.5 p-3.5 bg-[#0d0d0d] border border-[#1a1a1a] hover:border-gray-700 hover:text-white rounded-xl active:scale-95 transition-all text-gray-400 cursor-pointer select-none">
          <svg class="w-4 h-4 text-gray-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 1121.21 15H19" /></svg>
          <span class="text-[9px] font-mono tracking-widest uppercase">REPETIR</span>
        </button>
      </div>

      <!-- PLAYLIST EXERCISES (DAY'S AGENDA) -->
      <div class="flex-1 flex flex-col min-h-[160px]">
        <h3 class="text-[10px] font-mono text-gray-500 uppercase tracking-[0.25em] mb-2 flex items-center gap-1.5 select-none">
          <svg class="w-4 h-4 text-gray-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4 6h16M4 10h16M4 14h16M4 18h16" /></svg>
          SECUENCIA DE RUTINA
        </h3>
        
        <div id="exerciseList" class="flex-1 overflow-y-auto max-h-[180px] custom-scroll space-y-1 pr-1">
          <!-- Injected Dynamically -->
        </div>
      </div>
      
    </main>

    <!-- FOOTER / UTILS -->
    <footer id="app_footer_brand" class="mt-4 pt-3 border-t border-[#1a1a1a] text-center select-none flex flex-col items-center gap-1">
      <div class="flex justify-between w-full text-[9px] tracking-widest uppercase text-gray-600 font-mono">
        <span>LOCAL_DB: ROUTINE_STANDALONE</span>
        <span>AUDIO: 440HZ_READY</span>
      </div>
    </footer>
  </div>

  <script>
    // Exact copy of initial sample routines so the file works stand-alone immediately!
    let routine = ${JSON.stringify(routine)};
    
    // Attempt load from standalone localStorage
    const saved = localStorage.getItem("personal_workout_routine");
    if (saved) {
      try {
        routine = JSON.parse(saved);
      } catch (e) {
        console.error("Local load failed", e);
      }
    }

    let selectedDay = 1;
    let currentIndex = 0;
    let timeLeft = 0;
    let isPlaying = false;
    
    // High-precision timing
    let startTime = 0;
    let timeLeftOnPause = 0;
    let timerInterval = null;

    // Helper: list available days
    function getAvailableDays() {
      const daysSet = new Set(routine.map(x => x.dia));
      return Array.from(daysSet).sort((a,b) => a - b);
    }

    function getActiveExercises() {
      return routine.filter(x => x.dia === selectedDay);
    }

    function getDayName(dia) {
      const dayNames = {
        1: "Lunes",
        2: "Martes",
        3: "Miércoles",
        4: "Jueves",
        5: "Viernes",
        6: "Sábado",
        7: "Domingo"
      };
      return dayNames[dia] || "Día " + dia;
    }

    // AudioContext Alarm Generator
    function playAlertSound() {
      try {
        const AudioContextClass = window.AudioContext || window.webkitAudioContext;
        if (!AudioContextClass) return;
        const ctx = new AudioContextClass();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = "sine";
        osc.frequency.setValueAtTime(440, ctx.currentTime);
        gain.gain.setValueAtTime(0.4, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.5);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        osc.stop(ctx.currentTime + 0.5);
      } catch (e) {
        console.warn("Audio Context beep barred", e);
      }
    }

    // Format utility MM:SS
    function formatTime(secs) {
      const m = Math.floor(secs / 60);
      const s = secs % 60;
      return String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0');
    }

    // Render the state of UI
    function render() {
      const activeExs = getActiveExercises();
      const currentEx = activeExs[currentIndex];
      
      // Selectors rendering
      const days = getAvailableDays();
      if (!days.includes(selectedDay) && days.length > 0) {
        selectedDay = days[0];
      }
      
      const daySelectorContainer = document.getElementById("daySelectorContainer");
      daySelectorContainer.innerHTML = "";
      days.forEach(day => {
        const isActive = day === selectedDay;
        const button = document.createElement("button");
        button.className = "px-3 py-2 text-[11px] font-mono tracking-widest uppercase border transition-all duration-200 cursor-pointer " + 
          (isActive 
            ? "bg-[#f97316] text-black font-bold border-[#f97316] shadow-md" 
            : "bg-[#1a1a1a] border-[#333] hover:border-gray-500 text-gray-400 hover:text-white");
        button.innerText = getDayName(day).toUpperCase();
        button.onclick = () => {
          selectedDay = day;
          currentIndex = 0;
          isPlaying = false;
          clearInterval(timerInterval);
          timerInterval = null;
          const currentExs = getActiveExercises();
          if (currentExs[0]) {
            timeLeft = currentExs[0].tiempo;
            timeLeftOnPause = currentExs[0].tiempo;
          }
          render();
        };
        daySelectorContainer.appendChild(button);
      });

      // Update standalone subtitle header
      const headerDayDisplay = document.getElementById("headerDayDisplay");
      if (headerDayDisplay) {
        headerDayDisplay.innerHTML = getDayName(selectedDay) + " <span class=\"text-[#f97316] font-mono\">• " + activeExs.length + " EJ</span>";
      }
      
      // Current exercise card
      if (currentEx) {
        document.getElementById("exerciseTitle").innerText = currentEx.ejercicio;
        document.getElementById("countdownDisplay").innerText = formatTime(timeLeft);
        
        // Progress percentage calculation
        const total = currentEx.tiempo;
        const currentElapsed = total - timeLeft;
        const pct = total > 0 ? (currentElapsed / total) * 100 : 0;
        document.getElementById("progressPercent").innerText = Math.round(pct) + "%";

        const horizontalProgress = document.getElementById("horizontalProgress");
        if (horizontalProgress) {
          horizontalProgress.style.width = pct + "%";
        }
        
        // Circle SVG strokeoffset: 276.46 is full radius (44 * 2 * PI)
        const circumference = 276.46;
        const offset = circumference - (pct / 100) * circumference;
        document.getElementById("progressCircle").setAttribute("stroke-dashoffset", offset);

        // Styling according to if exercise is a "Descanso" or Rest
        const isRest = currentEx.ejercicio.toLowerCase().includes("descanso") || currentEx.ejercicio.toLowerCase().includes("rest");
        const statusBadge = document.getElementById("statusBadge");
        const standaloneStatusDot = document.getElementById("standaloneStatusDot");
        
        if (isRest) {
          statusBadge.innerText = "DESCANSO";
          statusBadge.className = "text-[9px] tracking-[0.2em] font-mono uppercase text-amber-500";
          if (standaloneStatusDot) {
            standaloneStatusDot.className = "w-2 h-2 rounded-full bg-amber-500 animate-pulse";
          }
          document.getElementById("progressCircle").setAttribute("stroke", "#f59e0b"); // Amber
          if (horizontalProgress) {
            horizontalProgress.className = "h-full bg-amber-500 transition-all duration-300";
          }
        } else {
          statusBadge.innerText = "EJERCICIO";
          statusBadge.className = "text-[9px] tracking-[0.2em] font-mono uppercase text-orange-500";
          if (standaloneStatusDot) {
            standaloneStatusDot.className = "w-2 h-2 rounded-full bg-orange-500 animate-pulse";
          }
          document.getElementById("progressCircle").setAttribute("stroke", "#f97316"); // Orange
          if (horizontalProgress) {
            horizontalProgress.className = "h-full bg-orange-500 transition-all duration-300";
          }
        }

        // Siguiente exercise
        const nextEx = activeExs[currentIndex + 1];
        if (nextEx) {
          document.getElementById("nextExerciseDisplay").innerText = nextEx.ejercicio.toUpperCase() + " (" + formatTime(nextEx.tiempo) + ")";
        } else {
          document.getElementById("nextExerciseDisplay").innerText = "FIN DE LA SESIÓN 🎉";
        }
      } else {
        document.getElementById("exerciseTitle").innerText = "Sube tu rutina";
        document.getElementById("countdownDisplay").innerText = "00:00";
        document.getElementById("nextExerciseDisplay").innerText = "Ninguno";
      }

      // Controller Play/Pause icons toggle
      if (isPlaying) {
        document.getElementById("playIconSvg").classList.add("hidden");
        document.getElementById("pauseIconSvg").classList.remove("hidden");
      } else {
        document.getElementById("playIconSvg").classList.remove("hidden");
        document.getElementById("pauseIconSvg").classList.add("hidden");
      }

      // Exercises playlist rendering
      const listContainer = document.getElementById("exerciseList");
      listContainer.innerHTML = "";
      
      activeExs.forEach((item, idx) => {
        const isActive = idx === currentIndex;
        const isPast = idx < currentIndex;

        const div = document.createElement("div");
        div.className = "flex items-center justify-between p-3.5 border transition-all cursor-pointer " +
          (isActive 
            ? "bg-[#1a130f] border-[#f97316] border-l-4" 
            : isPast 
              ? "bg-[#0d0d0d]/40 border-[#1a1a1a] opacity-35" 
              : "bg-[#0d0d0d] border-[#1a1a1a] hover:bg-[#111]");

        div.onclick = () => {
          currentIndex = idx;
          isPlaying = false;
          clearInterval(timerInterval);
          timerInterval = null;
          timeLeft = item.tiempo;
          timeLeftOnPause = item.tiempo;
          render();
        };

        const leftSide = document.createElement("div");
        leftSide.className = "flex items-center gap-2.5 truncate pr-2";
        
        const idxBadge = document.createElement("span");
        idxBadge.className = "text-[10px] font-mono " + (isActive ? "text-[#f97316]" : "text-gray-600");
        idxBadge.innerText = String(idx + 1).padStart(2, '0');
        leftSide.appendChild(idxBadge);

        const textWrapper = document.createElement("div");
        textWrapper.className = "truncate";
        const nameText = document.createElement("p");
        nameText.className = "text-xs font-medium truncate " + (isActive ? "text-white" : "text-gray-400");
        nameText.innerText = item.ejercicio;
        textWrapper.appendChild(nameText);
        leftSide.appendChild(textWrapper);

        const rightSide = document.createElement("div");
        rightSide.className = "flex items-center gap-2 shrink-0";
        const timeBadge = document.createElement("span");
        timeBadge.className = "text-[10px] font-mono " + (isActive ? "text-orange-400 font-bold" : "text-gray-500");
        timeBadge.innerText = formatTime(item.tiempo);
        rightSide.appendChild(timeBadge);

        if (isActive) {
          const activePulse = document.createElement("div");
          activePulse.className = "w-1.5 h-1.5 rounded-full bg-orange-500 animate-pulse";
          rightSide.appendChild(activePulse);
        }

        div.appendChild(leftSide);
        div.appendChild(rightSide);
        listContainer.appendChild(div);
      });
    }

    // Interval high precision timing logic
    function toggleTimer() {
      const activeExs = getActiveExercises();
      if (activeExs.length === 0) return;

      isPlaying = !isPlaying;
      
      if (isPlaying) {
        startTime = Date.now();
        timeLeftOnPause = timeLeft;

        timerInterval = setInterval(() => {
          const elapsed = Date.now() - startTime;
          const elapsedSecs = Math.floor(elapsed / 1000);
          const nextTimeLeft = Math.max(0, timeLeftOnPause - elapsedSecs);

          if (nextTimeLeft !== timeLeft) {
            timeLeft = nextTimeLeft;
            render();
          }

          if (nextTimeLeft === 0) {
            clearInterval(timerInterval);
            timerInterval = null;
            playAlertSound();
            
            // Advance automatically
            const currentExs = getActiveExercises();
            if (currentIndex < currentExs.length - 1) {
              currentIndex++;
              timeLeft = currentExs[currentIndex].tiempo;
              timeLeftOnPause = currentExs[currentIndex].tiempo;
              isPlaying = true;
              startTime = Date.now();
              // Restart interval immediately for the next
              toggleTimer(); // Toggles off then we toggle it back on
              isPlaying = false; // dummy toggling correction
              toggleTimer();
            } else {
              isPlaying = false;
              currentIndex = 0;
              timeLeft = currentExs[0].tiempo;
              render();
              alert("¡Enhorabuena! Has completado el entrenamiento de hoy. 💪🎉");
            }
          }
        }, 100);
      } else {
        clearInterval(timerInterval);
        timerInterval = null;
      }
      render();
    }

    // Play/Pause button binder
    document.getElementById("btnPlayPause").onclick = toggleTimer;

    // Reset current exercise
    document.getElementById("btnRepeatExercise").onclick = () => {
      const activeExs = getActiveExercises();
      const currentEx = activeExs[currentIndex];
      if (!currentEx) return;
      
      isPlaying = false;
      clearInterval(timerInterval);
      timerInterval = null;
      timeLeft = currentEx.tiempo;
      timeLeftOnPause = currentEx.tiempo;
      render();
    };

    // Reset current day entirely
    document.getElementById("btnResetDay").onclick = () => {
      const activeExs = getActiveExercises();
      isPlaying = false;
      clearInterval(timerInterval);
      timerInterval = null;
      currentIndex = 0;
      if (activeExs[0]) {
        timeLeft = activeExs[0].tiempo;
        timeLeftOnPause = activeExs[0].tiempo;
      }
      render();
    };

    // Excel Parsing Binder
    document.getElementById("excelFile").onchange = (e) => {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = event.target.result;
          const workbook = XLSX.read(data, { type: "binary" });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet);
          
          if (json.length === 0) {
            alert("El archivo Excel está vacío.");
            return;
          }

          const items = [];
          for (let i = 0; i < json.length; i++) {
            const row = json[i];
            
            const norm = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
            const keys = Object.keys(row);
            
            const dayKey = keys.find(k => norm(k) === "dia" || norm(k) === "day");
            const nameKey = keys.find(k => norm(k) === "ejercicio" || norm(k) === "exercise" || norm(k) === "name");
            const timeKey = keys.find(k => norm(k) === "tiempo" || norm(k) === "time" || norm(k) === "duration");

            if (!dayKey || !nameKey || !timeKey) continue;

            const dayVal = parseInt(row[dayKey]);
            if (isNaN(dayVal) || dayVal < 1 || dayVal > 7) continue;

            const nameVal = String(row[nameKey] || "").trim();
            if (!nameVal) continue;

            const rawTime = row[timeKey];
            let secs = 0;
            let timeStr = "";

            if (typeof rawTime === 'number') {
              if (rawTime > 0 && rawTime < 1) {
                secs = Math.round(rawTime * 86400);
                const m = Math.floor(secs / 60);
                const s = secs % 60;
                timeStr = String(m).padStart(2, '0') + ":" + String(s).padStart(2, '0');
              } else {
                secs = Math.round(rawTime);
                const m = Math.floor(secs / 60);
                const s = secs % 60;
                timeStr = String(m).padStart(2, '0') + ":" + String(s).padStart(2, '0');
              }
            } else if (typeof rawTime === 'string') {
              const tStr = rawTime.trim();
              timeStr = tStr;
              if (tStr.includes(":")) {
                const pts = tStr.split(":");
                if (pts.length === 2) {
                  secs = parseInt(pts[0])*60 + parseInt(pts[1]);
                } else if (pts.length === 3) {
                  secs = parseInt(pts[0])*3600 + parseInt(pts[1])*60 + parseInt(pts[2]);
                }
              } else {
                secs = parseInt(tStr) || 30;
              }
            }

            if (secs <= 0) secs = 30;

            items.push({
              dia: dayVal,
              ejercicio: nameVal,
              tiempo: secs,
              tiempoOriginal: timeStr || secs + "s"
            });
          }

          if (items.length === 0) {
            alert("No se pudieron extraer filas válidas de 'Día', 'Ejercicio' y 'Tiempo'.");
            return;
          }

          routine = items;
          localStorage.setItem("personal_workout_routine", JSON.stringify(items));
          
          const days = getAvailableDays();
          selectedDay = days[0] || 1;
          currentIndex = 0;
          isPlaying = false;
          clearInterval(timerInterval);
          timerInterval = null;
          
          const activeExs = getActiveExercises();
          if (activeExs[0]) {
            timeLeft = activeExs[0].tiempo;
            timeLeftOnPause = activeExs[0].tiempo;
          }
          render();
          alert("¡Rutina Excel cargada con éxito! 💪 " + items.length + " ejercicios importados.");
        } catch (err) {
          alert("Error procesando Excel: " + err.message);
        }
      };
      reader.readAsBinaryString(file);
    };

    // Auto unlock audio context on initial body touch/click
    document.body.addEventListener("click", function unlock() {
      try {
        const AudioContextClass = window.AudioContext || window.webkitAudioContext;
        const ctx = new AudioContextClass();
        ctx.resume();
      } catch (e) {}
      document.body.removeEventListener("click", unlock);
    }, { once: true });

    // Initial setup
    const initialExs = getActiveExercises();
    if (initialExs[0]) {
      timeLeft = initialExs[0].tiempo;
      timeLeftOnPause = initialExs[0].tiempo;
    }
    render();
  </script>
</body>
</html>`;

    // Download compiled standalone single HTML file
    const blob = new Blob([compiledHTML], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "ProTimer_Entrenamiento.html";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Helper formatting seconds to MM:SS
  const formatTime = (seconds: number) => {
    const minutes = Math.floor(seconds / 60);
    const secs = seconds % 60;
    return `${String(minutes).padStart(2, "0")}:${String(secs).padStart(2, "0")}`;
  };

  // Compute progress of the active exercise
  const totalDuration = currentExercise ? currentExercise.tiempo : 0;
  const elapsedSeconds = totalDuration - timeLeft;
  const progressPercent = totalDuration > 0 ? (elapsedSeconds / totalDuration) * 100 : 0;

  // Next exercise helper text
  const nextExercise = activeExercises[currentIndex + 1];

  return (
    <div id="workout_app_container" className="min-h-screen bg-[#0a0a0a] text-white flex items-center justify-center font-sans antialiased selection:bg-orange-500 selection:text-black p-0 md:p-4">
      
      {/* Main viewport mimicking a premium native cell phone layout with Geometric Balance theme */}
      <div id="phone_wrapper_card" className="w-full max-w-md bg-[#0a0a0a] md:rounded-2xl md:border-8 md:border-[#1a1a1a] md:shadow-2xl md:shadow-black/80 min-h-screen md:min-h-[820px] flex flex-col justify-between p-6 relative overflow-hidden">
        
        {/* TOP STATUS BAR ACCENT */}
        <div className="hidden md:flex justify-between items-center px-1 py-1 text-gray-500 text-[10px] font-mono select-none mb-2">
          <span>● INTENSITY CLOCK</span>
          <span>SYSTEM: PRECISE-V2</span>
          <span>OFFLINE LOCAL_DB</span>
        </div>

        {/* APP HEADER */}
        <header id="app_header" className="flex items-center justify-between border-b border-[#1a1a1a] pb-4 pt-1">
          <div>
            <h1 className="text-xs tracking-[0.3em] font-bold text-gray-500 uppercase">PROTIMER / PROGRESS</h1>
            <p className="text-xl font-light tracking-tight text-white">
              {getDayName(selectedDay)} <span className="text-orange-500 font-mono">• {activeExercises.length} EJ</span>
            </p>
          </div>

          <div className="flex items-center gap-2">
            {/* Download single-file standalone PWA */}
            <button
              onClick={handleExportSingleHTML}
              id="btn_export_html"
              title="Descargar archivo HTML Único para tu Celular"
              className="px-3 py-1.5 bg-[#1a1a1a] border border-[#333] hover:border-orange-500 rounded-md text-[9px] tracking-widest uppercase hover:bg-[#222] text-slate-300 hover:text-white transition-all cursor-pointer flex items-center gap-1 font-mono"
            >
              <Share2 className="w-3 h-3 text-orange-500" />
              <span>HTML Único</span>
            </button>
          </div>
        </header>

        {/* WORKOUT INTERFACE WORKSPACE */}
        <div id="workout_workspace" className="flex-1 flex flex-col justify-start gap-4 mt-4">
          
          {/* DAY SCROLLBAR PILLS */}
          <div id="day_pills_group" className="flex items-center gap-1.5 overflow-x-auto pb-1 no-scrollbar select-none">
            {availableDays.map((day) => {
              const isActive = day === selectedDay;
              return (
                <button
                  key={day}
                  id={`day_pill_${day}`}
                  onClick={() => handleSelectDay(day)}
                  className={`px-3 py-2 text-[11px] font-mono tracking-widest uppercase border transition-all duration-200 cursor-pointer ${
                    isActive
                      ? "bg-orange-500 text-black font-bold border-orange-500 shadow-md"
                      : "bg-[#1a1a1a] border-[#333] hover:border-gray-500 text-gray-400 hover:text-white"
                  }`}
                >
                  {getDayName(day).toUpperCase()}
                </button>
              );
            })}
          </div>

          {/* ACTIVE EXERCISE DISPLAY VIEWPORT */}
          <div
            id="active_timer_card"
            className="relative bg-[#0d0d0d] border border-[#1a1a1a] rounded-xl p-6 shadow-xl shadow-black/40 overflow-hidden flex flex-col items-center justify-center min-h-[320px]"
          >
            {/* Internal Clock Status Tag */}
            <div id="badge_status" className="absolute top-4 left-4 flex items-center gap-2">
              <div className={`w-2 h-2 rounded-full ${isPlaying ? "bg-orange-500 animate-pulse" : "bg-gray-700"}`}></div>
              <span className="text-[9px] tracking-[0.2em] font-mono uppercase text-gray-500">
                {currentExercise?.ejercicio.toLowerCase().includes("descanso") ? "RECOVERY / DESCANSO" : "ACTIVE / EJERCICIO"}
              </span>
            </div>

            {/* Total exercises progress count */}
            <div className="absolute top-4 right-4 text-[10px] font-mono tracking-widest text-gray-500">
              {String(currentIndex + 1).padStart(2, '0')} / {String(activeExercises.length).padStart(2, '0')}
            </div>

            {/* Big Name of active exercise */}
            <h2 id="current_exercise_name" className="text-2xl md:text-3xl font-bold tracking-tighter text-center mt-6 text-white uppercase line-clamp-2 px-2">
              {currentExercise ? currentExercise.ejercicio : "Sube tu rutina Excel"}
            </h2>

            {/* MASSIVE TICKING CLOCK & SVG RING */}
            <div className="relative w-52 h-52 md:w-56 md:h-56 my-3 flex items-center justify-center">
              <svg className="w-full h-full -rotate-90" viewBox="0 0 100 100">
                {/* Background Ring */}
                <circle cx="50" cy="50" r="44" stroke="#1a1a1a" strokeWidth="6" fill="none" />
                {/* Progress Active Ring */}
                <circle
                  cx="50"
                  cy="50"
                  r="44"
                  stroke="#f97316"
                  strokeWidth="6"
                  fill="none"
                  strokeLinecap="round"
                  strokeDasharray="276.46"
                  strokeDashoffset={276.46 - (progressPercent / 100) * 276.46}
                  className="transition-all duration-300"
                />
              </svg>

              {/* Digital countdown */}
              <div className="absolute flex flex-col items-center justify-center">
                <span id="numeric_countdown" className="text-5xl font-mono font-bold text-white tracking-tighter">
                  {formatTime(timeLeft)}
                </span>
                <span className="text-[9px] text-gray-500 tracking-[0.25em] font-mono uppercase mt-1">
                  {isPlaying ? "RUNNING" : "PAUSED"}
                </span>
              </div>
            </div>

            {/* Progress percentage slider bar straight line */}
            <div className="w-full max-w-[220px] h-1 bg-[#1a1a1a] rounded-full overflow-hidden relative mb-4">
              <div
                className="h-full bg-orange-500 transition-all duration-300"
                style={{ width: `${progressPercent}%` }}
              ></div>
            </div>

            {/* PREVIEW: NEXT EXERCISE BAR */}
            <div id="next_exercise_panel" className="w-full bg-[#0a0a0a] border border-[#1a1a1a] rounded-lg py-2.5 px-3 flex items-center justify-center gap-1.5 text-center">
              <span className="text-[10px] text-gray-500 tracking-wider font-mono uppercase">SIGUIENTE:</span>
              <span id="next_exercise_preview" className="text-[11px] font-bold text-white truncate max-w-[200px]">
                {nextExercise ? `${nextExercise.ejercicio.toUpperCase()} (${formatTime(nextExercise.tiempo)})` : "FIN DE LA SESIÓN 🎉"}
              </span>
            </div>
          </div>

          {/* TOUCH CONTROLLER PANEL */}
          <div id="controller_buttons_panel" className="grid grid-cols-3 items-center gap-3">
            {/* RESET DAY (REINICIAR DÍA) */}
            <button
              id="btn_reset_day"
              onClick={handleResetDay}
              title="Volver al primer ejercicio"
              className="flex flex-col items-center justify-center gap-1.5 p-3.5 bg-[#0d0d0d] border border-[#1a1a1a] hover:border-gray-700 hover:text-white rounded-xl active:scale-95 transition-all text-gray-400 cursor-pointer select-none"
            >
              <RotateCcw className="w-4 h-4 text-gray-500" />
              <span className="text-[9px] font-mono tracking-widest uppercase">REINICIAR</span>
            </button>

            {/* START/PAUSE ACTION (MASSIVE CENTER PILL) */}
            <button
              id="btn_play_pause"
              onClick={handleTogglePlay}
              className={`h-16 w-16 mx-auto rounded-full flex items-center justify-center active:scale-95 transition-all cursor-pointer border-4 border-[#0a0a0a] select-none shadow-lg ${
                isPlaying
                  ? "bg-white text-black hover:bg-gray-200 shadow-white/10"
                  : "bg-[#f97316] text-black hover:bg-[#ff8c3b] shadow-orange-500/20"
              }`}
            >
              {isPlaying ? <Pause className="w-6 h-6 fill-current" /> : <Play className="w-6 h-6 fill-current ml-0.5" />}
            </button>

            {/* REPEAT ACTIVE EXERCISE (REPETIR) */}
            <button
              id="btn_repeat_exercise"
              onClick={handleRepeatExercise}
              title="Reiniciar este ejercicio"
              className="flex flex-col items-center justify-center gap-1.5 p-3.5 bg-[#0d0d0d] border border-[#1a1a1a] hover:border-gray-700 hover:text-white rounded-xl active:scale-95 transition-all text-gray-400 cursor-pointer select-none"
            >
              <RefreshCw className="w-4 h-4 text-gray-500" />
              <span className="text-[9px] font-mono tracking-widest uppercase">REPETIR</span>
            </button>
          </div>

          {/* DYNAMIC LIST OF TODAY'S EXERCISES */}
          <div id="playlist_section" className="flex-1 flex flex-col min-h-[160px]">
            <h3 className="text-[10px] font-mono text-gray-500 uppercase tracking-[0.25em] mb-2 flex items-center gap-1.5 select-none">
              <Calendar className="w-4 h-4 text-gray-600" />
              SECUENCIA DE RUTINA ({activeExercises.length})
            </h3>

            <div id="exercise_list_container" className="flex-1 overflow-y-auto max-h-[190px] pr-1 space-y-1 custom-scroll">
              {activeExercises.map((item, idx) => {
                const isActive = idx === currentIndex;
                const isPast = idx < currentIndex;

                return (
                  <div
                    key={`${item.dia}_${idx}_${item.ejercicio}`}
                    id={`list_item_${idx}`}
                    onClick={() => handleSelectExerciseDirectly(idx)}
                    className={`flex items-center justify-between p-3.5 border transition-all cursor-pointer ${
                      isActive
                        ? "bg-[#1a130f] border-orange-500 border-l-4"
                        : isPast
                          ? "bg-[#0d0d0d]/40 border-[#1a1a1a] opacity-35 hover:opacity-70"
                          : "bg-[#0d0d0d] border-[#1a1a1a] hover:bg-[#111]"
                    }`}
                  >
                    <div className="flex items-center gap-2.5 truncate pr-2">
                      <span
                        className={`text-[10px] font-mono ${
                          isActive
                            ? "text-orange-500"
                            : "text-gray-600"
                        }`}
                      >
                        {String(idx + 1).padStart(2, '0')}
                      </span>
                      <p className={`text-xs font-medium truncate ${isActive ? "text-white" : "text-gray-400"}`}>
                        {item.ejercicio}
                      </p>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <span
                        className={`text-[10px] font-mono ${
                          isActive ? "text-orange-400 font-bold" : "text-gray-500"
                        }`}
                      >
                        {formatTime(item.tiempo)}
                      </span>
                      {isActive && <div className="w-1.5 h-1.5 rounded-full bg-orange-500 animate-pulse"></div>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* EXCEL CONFIGURATION BOX */}
          <div id="excel_upload_drawer" className="bg-[#0d0d0d] border border-[#1a1a1a] rounded-xl p-4 mt-1 flex flex-col gap-3">
            <div className="flex items-start gap-2.5">
              <FileSpreadsheet className="w-5 h-5 text-orange-500 shrink-0 mt-0.5" />
              <div>
                <h4 className="text-[11px] font-mono text-white uppercase tracking-wider">DATOS LOCALES & EXCEL</h4>
                <p className="text-[11px] text-gray-500 mt-0.5 leading-relaxed">
                  Sube un Excel con columnas <span className="text-white">Día</span> (1-7), <span className="text-white">Ejercicio</span> y <span className="text-white">Tiempo</span>. Los datos persisten de manera local.
                </p>
              </div>
            </div>

            {/* BUTTON CONTROLS FOR EXCEL */}
            <div className="grid grid-cols-2 gap-2 mt-1">
              <label className="flex items-center justify-center gap-1.5 py-2 px-3 bg-[#1a1a1a] border border-[#333] hover:border-orange-500 rounded-md text-xs font-bold text-slate-300 hover:text-white transition-all cursor-pointer">
                <Upload className="w-3.5 h-3.5 text-orange-500" />
                <span className="tracking-wide uppercase text-[10px] font-mono">IMPORTAR XLSX</span>
                <input type="file" onChange={handleFileUpload} accept=".xlsx, .xls" className="hidden" />
              </label>

              <button
                onClick={handleDownloadSampleExcel}
                id="btn_download_sample"
                className="flex items-center justify-center gap-1.5 py-2 px-3 bg-[#1a1a1a] border border-[#333] hover:border-gray-500 rounded-md text-xs font-bold text-slate-300 hover:text-white transition-all cursor-pointer"
              >
                <Download className="w-3.5 h-3.5 text-gray-400" />
                <span className="tracking-wide uppercase text-[10px] font-mono">PLANTILLA</span>
              </button>
            </div>

            {/* SUCCESS AND ERROR NOTIFICATIONS */}
            {excelError && (
              <div id="excel_error_banner" className="p-2.5 bg-red-500/10 border border-red-500/20 text-red-400 rounded-lg text-2xs font-semibold flex items-start gap-1.5 leading-relaxed">
                <Info className="w-3.5 h-3.5 shrink-0 text-red-400 mt-0.5" />
                <span>{excelError}</span>
              </div>
            )}

            {excelSuccess && (
              <div id="excel_success_banner" className="p-2.5 bg-orange-500/10 border border-orange-500/20 text-orange-400 rounded-lg text-2xs font-semibold flex items-center gap-1.5">
                <CheckCircle2 className="w-3.5 h-3.5 text-orange-500" />
                <span>¡Rutina cargada con éxito!</span>
              </div>
            )}

            {/* Reset to sample */}
            <button
              onClick={handleResetToDefault}
              id="btn_restore_routine"
              className="text-[9px] text-gray-600 hover:text-orange-400/80 font-bold uppercase tracking-widest text-center mt-1 self-center transition-all cursor-pointer font-mono"
            >
              RESTABLECER PRUEBA
            </button>
          </div>

        </div>

        {/* BOTTOM BRANDING */}
        <footer id="app_footer_brand" className="mt-4 pt-3 border-t border-[#1a1a1a] text-center select-none flex flex-col items-center gap-1">
          <div className="flex justify-between w-full text-[9px] tracking-widest uppercase text-gray-600 font-mono">
            <span>LOCAL_DB: ROUTINE_ACTV</span>
            <span>AUDIO: 440HZ_READY</span>
          </div>
        </footer>

      </div>
    </div>
  );
}
