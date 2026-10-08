/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Warehouse, Layers, Download, Plus, Trash2, CheckCircle2, 
  AlertTriangle, Settings, FileSpreadsheet, FileText, ArrowRight, ShieldAlert, 
  Sliders, Upload, Map, Search, Eye, Grid3X3, RefreshCw, ChevronRight, HelpCircle
} from 'lucide-react';

interface PassConfig {
  id: string;
  on: boolean;
  leftRack: string;
  rightRack: string;
  fromBay: number;
  toBay: number;
  levelDirection: 'up' | 'down';
  bayFilter: 'all' | 'even' | 'odd';
}

interface AisleConfig {
  id: string;
  name: string;
  bays: number;
  levelCount: number;
  levelsString: string;
  positionsPerBay: number;
  altBays: string;
  altPositionsPerBay: number;
  aisleSide: 'L' | 'R' | '';
  lockedPositions: string; // e.g. "3" or "2,4"
}

interface LocationRecord {
  id: string;
  aisle: string;
  bay: number;
  position: number;
  level: string;
  side: string;
  isLocked: boolean;
  lockStatus: string;
  pickSeq: number;
  putSeq: number;
  checkSeq: number;
  zone: string;
  workZone: string;
  locType: string;
  siteId: string;
  source: 'builder' | 'csv';
}

const LEVEL_LETTERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];
const LEVEL_COLORS: Record<string, string> = {
  A: '#38bdf8', // sky
  B: '#a855f7', // purple
  C: '#22c55e', // green
  D: '#f97316', // orange
  E: '#ec4899', // pink
  F: '#eab308', // yellow
  G: '#06b6d4', // cyan
  H: '#f43f5e'  // rose
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'builder' | 'map' | 'sequence' | 'passes' | 'export'>('builder');
  const [passes, setPasses] = useState<PassConfig[]>([]);

  // --- Dynamic Aisle Builder State ---
  const [siteId, setSiteId] = useState('DC4');
  const [globalLockedPos, setGlobalLockedPos] = useState('3');
  const [defaultLocType, setDefaultLocType] = useState('Tag FIFO');
  const [defaultZone, setDefaultZone] = useState('PICK');

  // Bulk Generator State
  const [bulkPrefix, setBulkPrefix] = useState('');
  const [bulkStartNum, setBulkStartNum] = useState(1);
  const [bulkEndNum, setBulkEndNum] = useState(6);
  const [bulkBays, setBulkBays] = useState(18);
  const [bulkLevels, setBulkLevels] = useState(6);
  const [bulkPositions, setBulkPositions] = useState(3);
  const [bulkAltBays, setBulkAltBays] = useState('');
  const [bulkAltPositions, setBulkAltPositions] = useState(4);
  const [bulkLockRule, setBulkLockRule] = useState('3');

  // Aisle Configurations
  const [aisles, setAisles] = useState<AisleConfig[]>([
    { id: '1', name: 'A01', bays: 18, levelCount: 6, levelsString: 'A,B,C,D,E,F', positionsPerBay: 3, altBays: '', altPositionsPerBay: 4, aisleSide: 'L', lockedPositions: '3' }
  ]);

  // --- Sequence Engine Parameters ---
  const [seqStart, setSeqStart] = useState(5);
  const [seqStep, setSeqStep] = useState(5);
  const [corridorDirection, setCorridorDirection] = useState<'forward' | 'reverse'>('forward');
  const [pairTransition, setPairTransition] = useState<'uturn' | 'straight'>('uturn');

  // --- Master Locations State & CSV Import ---
  const [locations, setLocations] = useState<LocationRecord[]>([]);
  const [csvRawText, setCsvRawText] = useState('');
  const [csvStatus, setCsvStatus] = useState<{ type: 'ok' | 'err'; message: string } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // --- Map Visualizer State ---
  const [mapViewMode, setMapViewMode] = useState<'pick' | 'put' | 'check' | 'level' | 'lock'>('pick');
  const [searchFilter, setSearchFilter] = useState('');
  const [selectedLocationId, setSelectedLocationId] = useState<string | null>(null);

  // Parse position locking rules (e.g., "3", "2,4", "2-4")
  const parseLockRule = (rule: string, maxPos: number): number[] => {
    const lockedSet = new Set<number>();
    String(rule || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean)
      .forEach(part => {
        if (part.includes('-')) {
          const [s, e] = part.split('-').map(v => parseInt(v.trim(), 10));
          if (!isNaN(s) && !isNaN(e)) {
            for (let i = Math.min(s, e); i <= Math.max(s, e); i++) {
              if (i >= 1 && i <= maxPos) lockedSet.add(i);
            }
          }
        } else {
          const p = parseInt(part, 10);
          if (!isNaN(p) && p >= 1 && p <= maxPos) lockedSet.add(p);
        }
      });
    return Array.from(lockedSet);
  };

  // Add individual aisle
  const handleAddAisle = () => {
    const nextIdx = aisles.length + 1;
    const name = `A${String(nextIdx).padStart(2, '0')}`;
    setAisles([
      ...aisles,
      {
        id: Date.now().toString(),
        name,
        bays: 18,
        levelCount: 6,
        levelsString: 'A,B,C,D,E,F',
        positionsPerBay: 3,
        altBays: '',
        altPositionsPerBay: 4,
        aisleSide: nextIdx % 2 === 1 ? 'L' : 'R',
        lockedPositions: globalLockedPos
      }
    ]);
  };

  // Delete aisle
  const handleDeleteAisle = (id: string) => {
    setAisles(aisles.filter(a => a.id !== id));
  };

  // Update aisle config
  const handleUpdateAisle = (id: string, field: keyof AisleConfig, val: any) => {
    setAisles(aisles.map(a => {
      if (a.id !== id) return a;
      if (field === 'levelCount') {
        const count = Math.max(1, Math.min(12, parseInt(val, 10) || 1));
        const levels = LEVEL_LETTERS.slice(0, count).join(',');
        return { ...a, levelCount: count, levelsString: levels };
      }
      return { ...a, [field]: val };
    }));
  };

  // Batch generate aisles
  const handleBulkGenerateAisles = () => {
    if (bulkEndNum < bulkStartNum) {
      alert('End aisle number must be greater than or equal to start aisle number.');
      return;
    }
    const newAisles: AisleConfig[] = [];
    const countLevels = Math.max(1, Math.min(12, bulkLevels));
    const levelsStr = LEVEL_LETTERS.slice(0, countLevels).join(',');

    for (let i = bulkStartNum; i <= bulkEndNum; i++) {
      const name = `${bulkPrefix}${String(i).padStart(2, '0')}`;
      newAisles.push({
        id: `${Date.now()}_${i}`,
        name,
        bays: bulkBays,
        levelCount: countLevels,
        levelsString: levelsStr,
        positionsPerBay: bulkPositions,
        altBays: bulkAltBays,
        altPositionsPerBay: bulkAltPositions,
        aisleSide: (i % 2 === 1) ? 'L' : 'R',
        lockedPositions: bulkLockRule
      });
    }
    setAisles(newAisles);
    alert(`Generated ${newAisles.length} aisles! Click "Build Map & Sequences" to render.`);
  };

  // Build grid & compute Pick, Putaway, and Check sequences
  const handleBuildGridAndMap = () => {
    if (!aisles.length) {
      alert('Please define at least one aisle in the builder.');
      return;
    }

    const generated: LocationRecord[] = [];
    let currentPickSeq = seqStart;

    // 1. Group aisles into corridor pairs (e.g., Aisle 0 & 1, Aisle 2 & 3)
    const aislePairs: AisleConfig[][] = [];
    for (let i = 0; i < aisles.length; i += 2) {
      if (i + 1 < aisles.length) {
        aislePairs.push([aisles[i], aisles[i + 1]]);
      } else {
        aislePairs.push([aisles[i]]);
      }
    }

    // Traversal across corridor pairs
    aislePairs.forEach((pair, pairIdx) => {
      // Determine bay order (U-turn alternate or straight)
      const isReverseBay = pairTransition === 'uturn' && (pairIdx % 2 === 1);
      const maxBaysInPair = Math.max(...pair.map(a => a.bays));

      const bayIndices: number[] = [];
      for (let b = 1; b <= maxBaysInPair; b++) bayIndices.push(b);
      if (isReverseBay) bayIndices.reverse();

      bayIndices.forEach(bayNum => {
        // Collect all unique levels for this bay across both aisles in the pair
        const allLevelsSet = new Set<string>();
        pair.forEach(a => {
          if (bayNum <= a.bays) {
            a.levelsString.split(',').map(l => l.trim()).filter(Boolean).forEach(l => allLevelsSet.add(l));
          }
        });
        const allLevels = Array.from(allLevelsSet).sort(); // Sort levels e.g. A, B, C

        allLevels.forEach(level => {
          // Within bay & level: iterate through aisles in pair (Left, then Right)
          pair.forEach(aisle => {
            if (bayNum > aisle.bays) return;
            const aisleLevels = aisle.levelsString.split(',').map(l => l.trim()).filter(Boolean);
            if (!aisleLevels.includes(level)) return;

            const altBaysList = parseLockRule(aisle.altBays, aisle.bays);
            const currentPositionsPerBay = altBaysList.includes(bayNum) ? aisle.altPositionsPerBay : aisle.positionsPerBay;
            
            let seqWithinAisleBase = 0;
            for (let prev = 1; prev < bayNum; prev++) {
               seqWithinAisleBase += altBaysList.includes(prev) ? aisle.altPositionsPerBay : aisle.positionsPerBay;
            }

            const effectiveLockRule = aisle.lockedPositions.trim() || globalLockedPos;
            const lockedPosList = parseLockRule(effectiveLockRule, currentPositionsPerBay);

            for (let pos = 1; pos <= currentPositionsPerBay; pos++) {
              const isLocked = lockedPosList.includes(pos);

              // DLx Location ID code: [Aisle][3-digit sequential bay][Level]
              const seqWithinAisle = seqWithinAisleBase + pos;
              const bayCode = String(seqWithinAisle).padStart(3, '0');
              const id = `${aisle.name}${bayCode}${level}`;

              generated.push({
                id,
                aisle: aisle.name,
                bay: bayNum,
                position: pos,
                level,
                side: aisle.aisleSide || (pos % 2 === 1 ? 'L' : 'R'),
                isLocked,
                lockStatus: isLocked ? 'Locked' : 'Unlock',
                pickSeq: currentPickSeq,
                putSeq: 0,
                checkSeq: 0,
                zone: defaultZone,
                workZone: 'A1',
                locType: defaultLocType,
                siteId,
                source: 'builder'
              });
              currentPickSeq += seqStep;
            }
          });
        });
      });
    });

    // 2. Putaway Sequence: exact reverse traversal of Pick Sequence
    const reversed = [...generated].reverse();
    let currentPutSeq = seqStart;
    reversed.forEach(row => {
      row.putSeq = currentPutSeq;
      currentPutSeq += seqStep;
    });

    // 3. Check Sequence: Scissors-lift stock check (Level first, Bay 1 upward, per aisle)
    let currentCheckSeq = seqStart;
    aisles.forEach(aisle => {
      const levels = aisle.levelsString.split(',').map(l => l.trim()).filter(Boolean);
      const altBaysList = parseLockRule(aisle.altBays, aisle.bays);
      
      levels.forEach(level => {
        let seqWithinAisleBase = 0;
        for (let b = 1; b <= aisle.bays; b++) {
          const currentPositionsPerBay = altBaysList.includes(b) ? aisle.altPositionsPerBay : aisle.positionsPerBay;
          for (let p = 1; p <= currentPositionsPerBay; p++) {
            const seqWithinAisle = seqWithinAisleBase + p;
            const bayCode = String(seqWithinAisle).padStart(3, '0');
            const targetId = `${aisle.name}${bayCode}${level}`;

            const item = generated.find(r => r.id === targetId);
            if (item) {
              item.checkSeq = currentCheckSeq;
              currentCheckSeq += seqStep;
            }
          }
          seqWithinAisleBase += currentPositionsPerBay;
        }
      });
    });

    setLocations(generated);
    setActiveTab('map');
  };

  const handleResequenceLocations = () => {
    if (!locations.length) return;

    // 1. Identify unique aisles and sort them according to corridor direction
    const uniqueAislesList = Array.from(new Set(locations.map(l => l.aisle))).sort();
    if (corridorDirection === 'reverse') uniqueAislesList.reverse();
    
    // 2. Sort the locations array based on Pick Sequence Engine logic
    const sortedLocations = [...locations].sort((a, b) => {
      const aAisleIdx = uniqueAislesList.indexOf(a.aisle);
      const bAisleIdx = uniqueAislesList.indexOf(b.aisle);
      
      const aPairIdx = Math.floor(aAisleIdx / 2);
      const bPairIdx = Math.floor(bAisleIdx / 2);
      
      if (aPairIdx !== bPairIdx) return aPairIdx - bPairIdx;
      
      const isUTurn = pairTransition === 'uturn' && aPairIdx % 2 === 1;
      const aEffectiveBay = isUTurn ? 9999 - a.bay : a.bay;
      const bEffectiveBay = isUTurn ? 9999 - b.bay : b.bay;
      
      if (aEffectiveBay !== bEffectiveBay) return aEffectiveBay - bEffectiveBay;
      
      if (a.level !== b.level) return a.level.localeCompare(b.level);
      
      const aAisleInPair = aAisleIdx % 2;
      const bAisleInPair = bAisleIdx % 2;
      if (aAisleInPair !== bAisleInPair) return aAisleInPair - bAisleInPair;
      
      return a.position - b.position;
    });

    // 3. Re-assign Pick and Putaway sequences
    let currentPickSeq = seqStart;
    sortedLocations.forEach(loc => {
      loc.pickSeq = currentPickSeq;
      currentPickSeq += seqStep;
    });

    let currentPutSeq = seqStart;
    [...sortedLocations].reverse().forEach(loc => {
      loc.putSeq = currentPutSeq;
      currentPutSeq += seqStep;
    });

    // 4. Re-assign Check sequences (Aisle -> Level -> Bay -> Position)
    const checkSorted = [...locations].sort((a, b) => {
      const aAisleIdx = uniqueAislesList.indexOf(a.aisle);
      const bAisleIdx = uniqueAislesList.indexOf(b.aisle);
      if (aAisleIdx !== bAisleIdx) return aAisleIdx - bAisleIdx;
      
      if (a.level !== b.level) return a.level.localeCompare(b.level);
      if (a.bay !== b.bay) return a.bay - b.bay;
      return a.position - b.position;
    });

    let currentCheckSeq = seqStart;
    checkSorted.forEach(loc => {
      const target = sortedLocations.find(s => s.id === loc.id);
      if (target) {
        target.checkSeq = currentCheckSeq;
        currentCheckSeq += seqStep;
      }
    });

    setLocations(sortedLocations);
    setCsvStatus({ type: 'ok', message: `Successfully re-sequenced ${sortedLocations.length.toLocaleString()} imported locations!` });
    setActiveTab('map');
  };

  const applyPassesToLocations = () => {
    if (!locations.length) return;
    const sorted = [...locations].sort((a, b) => {
      const passIdxA = passes.findIndex(p => p.on && (p.leftRack === a.aisle || p.rightRack === a.aisle) && a.bay >= Math.min(p.fromBay, p.toBay) && a.bay <= Math.max(p.fromBay, p.toBay) && (p.bayFilter === 'all' || (p.bayFilter === 'even' && a.bay % 2 === 0) || (p.bayFilter === 'odd' && a.bay % 2 === 1)));
      const passIdxB = passes.findIndex(p => p.on && (p.leftRack === b.aisle || p.rightRack === b.aisle) && b.bay >= Math.min(p.fromBay, p.toBay) && b.bay <= Math.max(p.fromBay, p.toBay) && (p.bayFilter === 'all' || (p.bayFilter === 'even' && b.bay % 2 === 0) || (p.bayFilter === 'odd' && b.bay % 2 === 1)));
      
      if (passIdxA === -1 && passIdxB === -1) return 0;
      if (passIdxA === -1) return 1;
      if (passIdxB === -1) return -1;
      
      if (passIdxA !== passIdxB) return passIdxA - passIdxB;
      
      const pass = passes[passIdxA];
      if (a.bay !== b.bay) {
         const baySign = Math.sign(pass.toBay - pass.fromBay) || 1;
         return (a.bay - b.bay) * baySign;
      }
      
      if (a.level !== b.level) {
         const lvlSign = pass.levelDirection === 'up' ? 1 : -1;
         return a.level.localeCompare(b.level) * lvlSign;
      }
      
      const aIsLeft = a.aisle === pass.leftRack;
      const bIsLeft = b.aisle === pass.leftRack;
      if (aIsLeft !== bIsLeft) return aIsLeft ? -1 : 1;
      
      return a.position - b.position;
    });

    let pick = seqStart;
    sorted.forEach(l => { l.pickSeq = pick; pick += seqStep; });
    let put = seqStart;
    [...sorted].reverse().forEach(l => { l.putSeq = put; put += seqStep; });
    
    setLocations(sorted);
    setCsvStatus({ type: 'ok', message: `Applied Pass rules to ${sorted.length} locations!` });
    setActiveTab('map');
  };

  // CSV Import Parser
  const parseCSVData = (text: string) => {
    try {
      const lines = text.split(/[\r\n]+/).map(l => l.trim()).filter(Boolean);
      if (lines.length < 2) {
        setCsvStatus({ type: 'err', message: 'CSV file contains no records.' });
        return;
      }

      const headers = lines[0].split(/[,\t;]/).map(h => h.trim().toUpperCase());
      const locIdIdx = headers.findIndex(h => h.includes('LOCATION') || h === 'ID' || h === 'LOC');
      const aisleIdx = headers.findIndex(h => h === 'AISLE');
      const bayIdx = headers.findIndex(h => h === 'BAY');
      const posIdx = headers.findIndex(h => h === 'POSITION' || h === 'POS');
      const levelIdx = headers.findIndex(h => h === 'LEVEL');
      const sideIdx = headers.findIndex(h => h.includes('SIDE'));
      const lockIdx = headers.findIndex(h => h.includes('LOCK'));
      const pickIdx = headers.findIndex(h => h.includes('PICK'));
      const putIdx = headers.findIndex(h => h.includes('PUT'));
      const checkIdx = headers.findIndex(h => h.includes('CHECK'));

      const imported: LocationRecord[] = [];
      let seqCounter = seqStart;

      for (let i = 1; i < lines.length; i++) {
        const parts = lines[i].split(/[,\t;]/).map(p => p.trim());
        if (!parts.length || !parts[0]) continue;

        let locId = locIdIdx >= 0 ? parts[locIdIdx] : parts[0];
        locId = locId.toUpperCase();

        // Extract regex pattern if possible: e.g. A01004A -> Aisle A01, bayCode 004, Level A
        let aisle = aisleIdx >= 0 && parts[aisleIdx] ? parts[aisleIdx] : '';
        let bay = bayIdx >= 0 && parts[bayIdx] ? parseInt(parts[bayIdx], 10) : 1;
        let pos = posIdx >= 0 && parts[posIdx] ? parseInt(parts[posIdx], 10) : 1;
        let level = levelIdx >= 0 && parts[levelIdx] ? parts[levelIdx] : 'A';

        const match = locId.match(/^([A-Z0-9]{2,4})([0-9]{3})([A-Z])$/);
        if (match) {
          if (!aisle) aisle = match[1];
          const rawSeq = parseInt(match[2], 10);
          bay = Math.ceil(rawSeq / 3) || 1;
          pos = ((rawSeq - 1) % 3) + 1;
          level = match[3];
        } else if (!aisle) {
          aisle = locId.slice(0, 3) || 'A01';
        }

        const isLock = lockIdx >= 0 && /LOCK|BLACK|YES|Y|TRUE/i.test(parts[lockIdx]);

        imported.push({
          id: locId,
          aisle,
          bay: isNaN(bay) ? 1 : bay,
          position: isNaN(pos) ? 1 : pos,
          level,
          side: sideIdx >= 0 ? parts[sideIdx] : 'R',
          isLocked: isLock,
          lockStatus: isLock ? 'Locked' : 'Unlock',
          pickSeq: pickIdx >= 0 && parseInt(parts[pickIdx], 10) ? parseInt(parts[pickIdx], 10) : seqCounter,
          putSeq: putIdx >= 0 && parseInt(parts[putIdx], 10) ? parseInt(parts[putIdx], 10) : 0,
          checkSeq: checkIdx >= 0 && parseInt(parts[checkIdx], 10) ? parseInt(parts[checkIdx], 10) : 0,
          zone: defaultZone,
          workZone: 'A1',
          locType: defaultLocType,
          siteId,
          source: 'csv'
        });
        
        seqCounter += seqStep;
      }

      if (!imported.length) {
        setCsvStatus({ type: 'err', message: 'No valid locations recognized in CSV file.' });
        return;
      }

      setLocations(imported);
      
      const uniqueImportedAisles = Array.from(new Set(imported.map(l => l.aisle))).sort();
      if (passes.length === 0 && uniqueImportedAisles.length > 0) {
         const defaultPasses = [];
         for(let i=0; i<uniqueImportedAisles.length; i+=2) {
           defaultPasses.push({
             id: Date.now().toString() + i,
             on: true,
             leftRack: uniqueImportedAisles[i],
             rightRack: uniqueImportedAisles[i+1] || '',
             fromBay: 1,
             toBay: Math.max(...imported.filter(l => l.aisle === uniqueImportedAisles[i]).map(l => l.bay), 1),
             levelDirection: 'up' as const,
             bayFilter: 'all' as const
           });
         }
         setPasses(defaultPasses);
      }
      setCsvStatus({ type: 'ok', message: `Successfully parsed ${imported.length.toLocaleString()} locations. Please configure Passes.` });
      setActiveTab('passes');
    } catch (err: any) {
      setCsvStatus({ type: 'err', message: `Parse error: ${err.message}` });
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const content = ev.target?.result as string;
      setCsvRawText(content);
      parseCSVData(content);
    };
    reader.readAsText(file);
  };

  // Compute map matrix & metrics
  const uniqueAisles = useMemo(() => {
    return Array.from(new Set(locations.map(l => l.aisle))).sort();
  }, [locations]);

  const maxBayInMap = useMemo(() => {
    return locations.reduce((max, l) => Math.max(max, l.bay), 1);
  }, [locations]);

  const uniqueLevels = useMemo(() => {
    return Array.from(new Set(locations.map(l => l.level))).sort();
  }, [locations]);

  const maxPick = useMemo(() => locations.reduce((m, l) => Math.max(m, l.pickSeq), 0), [locations]);
  const maxPut = useMemo(() => locations.reduce((m, l) => Math.max(m, l.putSeq), 0), [locations]);
  const maxCheck = useMemo(() => locations.reduce((m, l) => Math.max(m, l.checkSeq), 0), [locations]);

  // Heatmap gradient color calculator
  const getHeatmapColor = (val: number, max: number) => {
    if (!val || !max) return '#1a1f2c';
    const ratio = Math.min(1, Math.max(0, val / max));
    const stops = [
      [139, 92, 246],  // violet-500
      [217, 70, 239],  // fuchsia-500
      [236, 72, 153],  // pink-500
      [244, 63, 94],   // rose-500
      [245, 158, 11]   // amber-500
    ];
    const segment = ratio * (stops.length - 1);
    const idx = Math.floor(segment);
    const t = segment - idx;
    if (idx >= stops.length - 1) return `rgb(${stops[stops.length - 1].join(',')})`;
    const r = Math.round(stops[idx][0] + (stops[idx + 1][0] - stops[idx][0]) * t);
    const g = Math.round(stops[idx][1] + (stops[idx + 1][1] - stops[idx][1]) * t);
    const b = Math.round(stops[idx][2] + (stops[idx + 1][2] - stops[idx][2]) * t);
    return `rgb(${r},${g},${b})`;
  };

  // Selected cell record
  const selectedCell = useMemo(() => {
    return locations.find(l => l.id === selectedLocationId) || null;
  }, [locations, selectedLocationId]);

  // Export to Excel (.xlsx)
  const exportToExcel = () => {
    if (!locations.length) {
      alert('No location data available. Build the grid or import a CSV first.');
      return;
    }
    const headers = [
      'PICK_SEQUENCE','PUT_SEQUENCE','CHECK_SEQUENCE','LOCATION_ID','AISLE','BAY','POSITION','LEVEL',
      'AISLE_SIDE','LOCK_STATUS','IS_LOCKED','ZONE_1','WORK_ZONE','LOC_TYPE','SITE_ID'
    ];
    const rows = locations.map(l => [
      l.pickSeq, l.putSeq, l.checkSeq, l.id, l.aisle, l.bay, l.position, l.level,
      l.side, l.lockStatus, l.isLocked ? 'Y' : 'N', l.zone, l.workZone, l.locType, l.siteId
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'INTERFACE_LOCATION');
    XLSX.writeFile(wb, `DC4_Warehouse_Locations_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Export to CSV
  const exportToCsv = () => {
    if (!locations.length) {
      alert('No location data available.');
      return;
    }
    const headers = 'PICK_SEQUENCE,PUT_SEQUENCE,CHECK_SEQUENCE,LOCATION_ID,AISLE,BAY,POSITION,LEVEL,AISLE_SIDE,LOCK_STATUS,IS_LOCKED,ZONE_1,WORK_ZONE,LOC_TYPE,SITE_ID';
    const lines = locations.map(l => 
      `${l.pickSeq},${l.putSeq},${l.checkSeq},${l.id},${l.aisle},${l.bay},${l.position},${l.level},${l.side},${l.lockStatus},${l.isLocked ? 'Y' : 'N'},${l.zone},${l.workZone},${l.locType},${l.siteId}`
    );
    const blob = new Blob([[headers, ...lines].join('\n')], { type: 'text/csv' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `DC4_Warehouse_Locations_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
  };

  return (
    <div className="flex flex-col h-screen bg-[#020617] text-slate-100 font-sans overflow-hidden">
      {/* Top Header */}
      <header className="px-6 py-3.5 bg-[#0f172a] border-b border-[#334155] flex justify-between items-center shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-violet-500/10 border border-violet-500/30 flex items-center justify-center text-violet-400">
            <Warehouse className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-sm font-extrabold tracking-tight">DC4 <span className="text-violet-400">DYNAMIC AISLE BUILDER</span> & MAP</h1>
            <p className="text-[10px] text-slate-400">RedPrairie DLx Dispatcher 2016.2 — Grid Architect & Sequence Visualizer</p>
          </div>
        </div>

        {/* Global Nav Tabs */}
        <div className="flex items-center gap-1.5 bg-[#0e1118] p-1 border border-[#334155] rounded-lg">
          <button
            onClick={() => setActiveTab('builder')}
            className={`px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 ${activeTab === 'builder' ? 'bg-violet-500 text-black' : 'text-slate-400 hover:text-white'}`}
          >
            <Grid3X3 className="w-3.5 h-3.5" /> Aisle Builder ({aisles.length})
          </button>
          <button
            onClick={() => setActiveTab('map')}
            className={`px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 ${activeTab === 'map' ? 'bg-violet-500 text-black' : 'text-slate-400 hover:text-white'}`}
          >
            <Map className="w-3.5 h-3.5" /> Warehouse Map ({locations.length.toLocaleString()})
          </button>
          <button
            onClick={() => setActiveTab('sequence')}
            className={`px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 ${activeTab === 'sequence' ? 'bg-violet-500 text-black' : 'text-slate-400 hover:text-white'}`}
          >
            <Sliders className="w-3.5 h-3.5" /> Sequence Rules
          </button>
          <button
            onClick={() => setActiveTab('passes')}
            className={`px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 ${activeTab === 'passes' ? 'bg-amber-500 text-black' : 'text-slate-400 hover:text-white'}`}
          >
            <RefreshCw className="w-3.5 h-3.5" /> Pass Planner
          </button>
          <button
            onClick={() => setActiveTab('export')}
            className={`px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 ${activeTab === 'export' ? 'bg-violet-500 text-black' : 'text-slate-400 hover:text-white'}`}
          >
            <Download className="w-3.5 h-3.5" /> DTO Export
          </button>
        </div>

        <div className="flex items-center gap-2">
          <button 
            onClick={exportToExcel}
            disabled={!locations.length}
            className="px-3.5 py-1.5 bg-fuchsia-500 hover:bg-fuchsia-400 disabled:opacity-40 text-black font-extrabold rounded-md text-xs transition flex items-center gap-1.5"
          >
            <Download className="w-3.5 h-3.5" /> Export .xlsx
          </button>
        </div>
      </header>

      {/* Main Content Area */}
      <div className="flex-1 overflow-hidden">
        
        {/* TAB 1: DYNAMIC AISLE BUILDER & CSV INGESTION */}
        {activeTab === 'builder' && (
          <div className="h-full grid grid-cols-[400px_1fr] overflow-hidden">
            {/* Left Control Column */}
            <div className="bg-[#0f172a] border-r border-[#334155] p-4 overflow-y-auto flex flex-col gap-4">
              
              {/* Batch Aisle Generator Card */}
              <div className="bg-[#1e293b] border border-[#334155] p-3.5 rounded-xl">
                <h2 className="text-[11px] uppercase font-bold tracking-wider text-violet-400 mb-2 flex items-center gap-2">
                  <Grid3X3 className="w-3.5 h-3.5" /> Quick Batch Row Generator
                </h2>
                <p className="text-[10px] text-slate-400 mb-3 leading-relaxed">
                  Quickly generate sequential aisles/rows with preset bay, level, and position counts.
                </p>
                <div className="grid grid-cols-3 gap-2 text-xs">
                  <div>
                    <label className="text-[9px] text-slate-400 font-semibold">Row Prefix</label>
                    <input type="text" value={bulkPrefix} onChange={(e) => setBulkPrefix(e.target.value)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white text-xs outline-none" />
                  </div>
                  <div>
                    <label className="text-[9px] text-slate-400 font-semibold">Start #</label>
                    <input type="number" value={bulkStartNum} onChange={(e) => setBulkStartNum(parseInt(e.target.value, 10) || 1)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white text-xs outline-none" min="1" />
                  </div>
                  <div>
                    <label className="text-[9px] text-slate-400 font-semibold">End #</label>
                    <input type="number" value={bulkEndNum} onChange={(e) => setBulkEndNum(parseInt(e.target.value, 10) || 1)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white text-xs outline-none" min="1" />
                  </div>
                  <div>
                    <label className="text-[9px] text-slate-400 font-semibold">Bay Count</label>
                    <input type="number" value={bulkBays} onChange={(e) => setBulkBays(parseInt(e.target.value, 10) || 1)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white text-xs outline-none" min="1" />
                  </div>
                  <div>
                    <label className="text-[9px] text-slate-400 font-semibold">Level Count</label>
                    <input type="number" value={bulkLevels} onChange={(e) => setBulkLevels(parseInt(e.target.value, 10) || 1)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white text-xs outline-none" min="1" max="12" />
                  </div>
                  <div>
                    <label className="text-[9px] text-slate-400 font-semibold">Positions/Bay</label>
                    <input type="number" value={bulkPositions} onChange={(e) => setBulkPositions(parseInt(e.target.value, 10) || 1)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white text-xs outline-none" min="1" />
                  </div>
                </div>

                <div className="mt-2.5 grid grid-cols-3 gap-2 text-xs">
                  <div>
                    <label className="text-[9px] text-violet-400 font-semibold">Alt Pos Bays</label>
                    <input type="text" value={bulkAltBays} onChange={(e) => setBulkAltBays(e.target.value)} placeholder="e.g. 1-5" className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-violet-400 text-xs outline-none" />
                    <span className="text-[9px] text-slate-500 mt-0.5 block">Bays with alt pos.</span>
                  </div>
                  <div>
                    <label className="text-[9px] text-violet-400 font-semibold">Alt Pos / Bay</label>
                    <input type="number" value={bulkAltPositions} onChange={(e) => setBulkAltPositions(parseInt(e.target.value, 10) || 1)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-violet-400 text-xs outline-none" min="1" />
                    <span className="text-[9px] text-slate-500 mt-0.5 block">Overrides count.</span>
                  </div>
                  <div>
                    <label className="text-[9px] text-amber-400 font-semibold">Locked Position Rule</label>
                    <input type="text" value={bulkLockRule} onChange={(e) => setBulkLockRule(e.target.value)} placeholder="e.g. 3 or 2,4" className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-amber-400 text-xs outline-none" />
                    <span className="text-[9px] text-slate-500 mt-0.5 block">e.g. "3" locks 3rd pos.</span>
                  </div>
                </div>

                <button 
                  onClick={handleBulkGenerateAisles}
                  className="w-full mt-3 py-2 bg-[#334155] hover:bg-slate-700 text-violet-400 font-bold rounded-lg text-xs transition flex items-center justify-center gap-1.5"
                >
                  <RefreshCw className="w-3.5 h-3.5" /> Apply Batch Rows ({bulkEndNum - bulkStartNum + 1} Aisles)
                </button>
              </div>

              {/* CSV Upload / Paste Card */}
              <div className="bg-[#1e293b] border border-[#334155] p-3.5 rounded-xl flex flex-col gap-2.5">
                <div className="flex justify-between items-center">
                  <h2 className="text-[11px] uppercase font-bold tracking-wider text-fuchsia-400 flex items-center gap-2">
                    <Upload className="w-3.5 h-3.5" /> Or Import CSV File
                  </h2>
                  <button 
                    onClick={() => fileInputRef.current?.click()}
                    className="px-2 py-1 bg-fuchsia-500/10 hover:bg-fuchsia-500/20 text-fuchsia-400 border border-fuchsia-500/30 rounded text-[10px] font-bold"
                  >
                    Upload File
                  </button>
                  <input 
                    type="file" 
                    ref={fileInputRef} 
                    onChange={handleFileUpload} 
                    accept=".csv,.txt,.tsv" 
                    className="hidden" 
                  />
                </div>
                <textarea 
                  value={csvRawText}
                  onChange={(e) => setCsvRawText(e.target.value)}
                  placeholder="Paste CSV records with headers:&#10;LOCATION_ID,AISLE,BAY,LEVEL,POSITION&#10;A01001A,A01,1,A,1&#10;A01002A,A01,1,A,2..."
                  className="w-full h-24 bg-black border border-[#334155] rounded p-2 text-[10px] font-mono text-fuchsia-400 outline-none resize-vertical"
                />
                <button 
                  onClick={() => parseCSVData(csvRawText)}
                  className="w-full py-1.5 bg-fuchsia-500 hover:bg-fuchsia-400 text-black font-bold rounded-lg text-xs transition"
                >
                  Parse & Map Pasted CSV
                </button>
                {csvStatus && (
                  <div className={`p-2 rounded text-[10px] font-mono ${csvStatus.type === 'err' ? 'bg-rose-500/10 text-rose-400 border border-rose-500/30' : 'bg-violet-500/10 text-violet-400 border border-violet-500/30'}`}>
                    {csvStatus.message}
                  </div>
                )}
                {locations.length > 0 && locations[0].source === 'csv' && (
                  <button 
                    onClick={handleResequenceLocations}
                    className="w-full py-1.5 bg-amber-500 hover:bg-amber-400 text-black font-bold rounded-lg text-xs transition mt-2"
                  >
                    Re-Sequence Imported Data using Rules
                  </button>
                )}
              </div>

              {/* Primary Action Button */}
              <button 
                onClick={handleBuildGridAndMap}
                className="w-full py-3 bg-violet-500 hover:bg-violet-400 text-black font-extrabold rounded-xl text-xs transition shadow-lg shadow-violet-500/10 flex items-center justify-center gap-2 mt-auto"
              >
                <Map className="w-4 h-4" /> Build Map & Calculate Sequences
              </button>
            </div>

            {/* Right Detailed Aisle Table */}
            <div className="p-5 bg-[#020617] overflow-auto flex flex-col gap-4">
              <div className="flex justify-between items-center">
                <div>
                  <h2 className="text-sm font-bold text-white">Defined Aisle & Row Specifications</h2>
                  <p className="text-xs text-slate-400">Configure bay counts, levels, positions per bay, and per-row lock overrides.</p>
                </div>
                <button 
                  onClick={handleAddAisle}
                  className="px-3 py-1.5 bg-violet-500/20 hover:bg-violet-500/30 text-violet-300 border border-violet-500/30 text-xs font-bold rounded-lg flex items-center gap-1.5 transition"
                >
                  <Plus className="w-3.5 h-3.5" /> Add Single Aisle
                </button>
              </div>

              {/* Editable Aisle List Grid */}
              <div className="border border-[#334155] rounded-xl overflow-hidden bg-[#1e293b]">
                <table className="w-full border-collapse font-mono text-xs text-left">
                  <thead>
                    <tr className="bg-[#0f172a] text-slate-400 border-b border-[#334155]">
                      <th className="p-3">Aisle / Row</th>
                      <th className="p-3">Bays</th>
                      <th className="p-3">Levels (Count)</th>
                      <th className="p-3">Positions / Bay</th>
                      <th className="p-3 text-violet-400" colSpan={2}>Alt Pos Bays & Count</th>
                      <th className="p-3">Side</th>
                      <th className="p-3 text-amber-400">Locked Pos Rule</th>
                      <th className="p-3 text-right">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#1e293b]">
                    {aisles.map((aisle) => (
                      <tr key={aisle.id} className="hover:bg-white/5 transition">
                        <td className="p-3">
                          <input 
                            type="text" 
                            value={aisle.name} 
                            onChange={(e) => handleUpdateAisle(aisle.id, 'name', e.target.value)}
                            className="w-24 bg-black border border-[#334155] focus:border-violet-500 rounded p-1 font-bold text-white text-xs outline-none"
                          />
                        </td>
                        <td className="p-3">
                          <input 
                            type="number" 
                            value={aisle.bays} 
                            onChange={(e) => handleUpdateAisle(aisle.id, 'bays', parseInt(e.target.value, 10) || 1)}
                            className="w-16 bg-black border border-[#334155] focus:border-violet-500 rounded p-1 text-white text-xs outline-none"
                            min="1"
                          />
                        </td>
                        <td className="p-3">
                          <div className="flex items-center gap-2">
                            <input 
                              type="number" 
                              value={aisle.levelCount} 
                              onChange={(e) => handleUpdateAisle(aisle.id, 'levelCount', e.target.value)}
                              className="w-14 bg-black border border-[#334155] focus:border-violet-500 rounded p-1 text-white text-xs outline-none"
                              min="1"
                              max="12"
                            />
                            <span className="text-[10px] text-slate-500 font-sans">({aisle.levelsString})</span>
                          </div>
                        </td>
                        <td className="p-3">
                          <input 
                            type="number" 
                            value={aisle.positionsPerBay} 
                            onChange={(e) => handleUpdateAisle(aisle.id, 'positionsPerBay', parseInt(e.target.value, 10) || 1)}
                            className="w-12 bg-black border border-[#334155] focus:border-violet-500 rounded p-1 text-white text-xs outline-none"
                            min="1"
                          />
                        </td>
                        <td className="p-1">
                          <input 
                            type="text" 
                            value={aisle.altBays} 
                            onChange={(e) => handleUpdateAisle(aisle.id, 'altBays', e.target.value)}
                            placeholder="e.g. 1-5"
                            className="w-16 bg-black border border-[#334155] focus:border-violet-400 rounded p-1 text-violet-400 text-xs outline-none font-bold"
                          />
                        </td>
                        <td className="p-1">
                          <input 
                            type="number" 
                            value={aisle.altPositionsPerBay} 
                            onChange={(e) => handleUpdateAisle(aisle.id, 'altPositionsPerBay', parseInt(e.target.value, 10) || 1)}
                            className="w-10 bg-black border border-[#334155] focus:border-violet-400 rounded p-1 text-violet-400 text-xs outline-none"
                            min="1"
                          />
                        </td>
                        <td className="p-3">
                          <select 
                            value={aisle.aisleSide} 
                            onChange={(e) => handleUpdateAisle(aisle.id, 'aisleSide', e.target.value)}
                            className="bg-black border border-[#334155] rounded p-1 text-white text-xs outline-none"
                          >
                            <option value="L">L</option>
                            <option value="R">R</option>
                          </select>
                        </td>
                        <td className="p-3">
                          <input 
                            type="text" 
                            value={aisle.lockedPositions} 
                            onChange={(e) => handleUpdateAisle(aisle.id, 'lockedPositions', e.target.value)}
                            placeholder="e.g. 3"
                            className="w-20 bg-black border border-[#334155] focus:border-amber-400 rounded p-1 text-amber-400 text-xs outline-none font-bold"
                          />
                        </td>
                        <td className="p-3 text-right">
                          {aisles.length > 1 && (
                            <button 
                              onClick={() => handleDeleteAisle(aisle.id)}
                              className="p-1 text-rose-400 hover:text-rose-300 transition"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Summary KPIs */}
              <div className="grid grid-cols-4 gap-3">
                <div className="bg-[#0f172a] border border-[#334155] p-3 rounded-xl text-center">
                  <div className="text-[10px] uppercase text-slate-400 font-semibold mb-1">Defined Aisles</div>
                  <div className="text-xl font-mono font-bold text-purple-400">{aisles.length}</div>
                </div>
                <div className="bg-[#0f172a] border border-[#334155] p-3 rounded-xl text-center">
                  <div className="text-[10px] uppercase text-slate-400 font-semibold mb-1">Estimated Locations</div>
                  <div className="text-xl font-mono font-bold text-violet-400">
                    {aisles.reduce((sum, a) => sum + (a.bays * a.levelCount * a.positionsPerBay), 0).toLocaleString()}
                  </div>
                </div>
                <div className="bg-[#0f172a] border border-[#334155] p-3 rounded-xl text-center">
                  <div className="text-[10px] uppercase text-slate-400 font-semibold mb-1">Max Bays Per Aisle</div>
                  <div className="text-xl font-mono font-bold text-fuchsia-400">
                    {Math.max(...aisles.map(a => a.bays), 0)}
                  </div>
                </div>
                <div className="bg-[#0f172a] border border-[#334155] p-3 rounded-xl text-center">
                  <div className="text-[10px] uppercase text-slate-400 font-semibold mb-1">Global Locked Pos Rule</div>
                  <div className="text-xl font-mono font-bold text-amber-400">{globalLockedPos || 'None'}</div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: INTERACTIVE WAREHOUSE MAP & HEATMAP */}
        {activeTab === 'map' && (
          <div className="h-full grid grid-cols-[1fr_320px] overflow-hidden">
            {/* Map Canvas & Navigation Controls */}
            <div className="p-4 bg-[#020617] overflow-auto flex flex-col gap-3">
              <div className="flex justify-between items-center shrink-0">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold text-slate-400">Heatmap Mode:</span>
                  <div className="flex bg-[#0f172a] border border-[#334155] p-0.5 rounded-lg">
                    {(['pick', 'put', 'check', 'level', 'lock'] as const).map(mode => (
                      <button
                        key={mode}
                        onClick={() => setMapViewMode(mode)}
                        className={`px-3 py-1 rounded text-xs font-bold transition ${mapViewMode === mode ? 'bg-violet-500 text-black' : 'text-slate-400 hover:text-white'}`}
                      >
                        {mode === 'pick' && 'Pick Seq'}
                        {mode === 'put' && 'Putaway Seq'}
                        {mode === 'check' && 'Check Seq'}
                        {mode === 'level' && 'Levels'}
                        {mode === 'lock' && 'Locked'}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="flex items-center gap-3">
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                    <input 
                      type="text" 
                      value={searchFilter}
                      onChange={(e) => setSearchFilter(e.target.value)}
                      placeholder="Filter location / aisle..."
                      className="pl-8 pr-3 py-1 bg-black border border-[#334155] rounded-md text-xs font-mono text-white outline-none w-48 focus:border-violet-500"
                    />
                  </div>
                  <span className="text-xs font-mono text-slate-400">
                    {locations.length.toLocaleString()} locations mapped
                  </span>
                </div>
              </div>

              {/* Heatmap Legend */}
              {(mapViewMode === 'pick' || mapViewMode === 'put' || mapViewMode === 'check') && (
                <div className="flex flex-col gap-1 shrink-0">
                  <div className="h-2 rounded bg-gradient-to-r from-violet-500 via-fuchsia-500 via-pink-500 via-rose-500 to-amber-500"></div>
                  <div className="flex justify-between text-[9px] font-mono text-slate-400">
                    <span>First Sequence ({seqStart + seqStep})</span>
                    <span>Max Sequence ({mapViewMode === 'put' ? maxPut : mapViewMode === 'check' ? maxCheck : maxPick})</span>
                  </div>
                </div>
              )}

              {/* 2D Interactive Grid Map Container */}
              <div className="flex-1 overflow-auto border border-[#334155] rounded-xl p-4 bg-[#0f172a]">
                {uniqueAisles.length === 0 ? (
                  <div className="h-full flex flex-col items-center justify-center text-slate-500 text-xs">
                    <Map className="w-8 h-8 mb-2 opacity-50" />
                    <span>No locations loaded. Go to the "Aisle Builder" tab to generate or import your grid.</span>
                  </div>
                ) : (
                  <div className="flex flex-col gap-5 min-w-max">
                    {uniqueAisles.map(aisleName => {
                      const aisleLocs = locations.filter(l => l.aisle === aisleName);
                      const aisleLevels = Array.from(new Set(aisleLocs.map(l => l.level))).sort().reverse();
                      const maxAisleBay = Math.max(...aisleLocs.map(l => l.bay), 1);

                      return (
                        <div key={aisleName} className="flex flex-col gap-1.5 p-3 rounded-lg border border-[#1e293b] bg-[#0c0e17]/80">
                          <div className="flex justify-between items-center text-xs font-mono pb-1 border-b border-[#1e293b]">
                            <span className="font-bold text-violet-400">Aisle {aisleName}</span>
                            <span className="text-[10px] text-slate-400">{aisleLocs.length} locations | {maxAisleBay} Bays</span>
                          </div>

                          {/* Levels & Bays Matrix */}
                          <div className="flex flex-col gap-1">
                            {aisleLevels.map(lvl => (
                              <div key={lvl} className="flex items-center gap-1">
                                <span className="w-6 text-[10px] font-mono text-slate-400 font-bold shrink-0 text-center">
                                  {lvl}
                                </span>
                                <div className="flex gap-0.5">
                                  {Array.from({ length: maxAisleBay }).map((_, bIdx) => {
                                    const bayNum = bIdx + 1;
                                    // Multiple positions in each bay
                                    const bayCells = aisleLocs.filter(l => l.bay === bayNum && l.level === lvl);
                                    if (!bayCells.length) {
                                      return <div key={bayNum} className="w-4 h-4 bg-black/40 rounded-sm"></div>;
                                    }

                                    return (
                                      <div key={bayNum} className="flex gap-px p-0.5 bg-black/60 rounded border border-[#334155]">
                                        {bayCells.sort((a, b) => a.position - b.position).map(cell => {
                                          let bg = '#1a1e2a';
                                          const q = searchFilter.trim().toUpperCase();
                                          const matchesSearch = q && (cell.id.includes(q) || cell.aisle.includes(q));

                                          if (mapViewMode === 'level') {
                                            bg = LEVEL_COLORS[cell.level] || '#38bdf8';
                                          } else if (mapViewMode === 'lock') {
                                            bg = cell.isLocked ? '#f59e0b' : '#1e293b';
                                          } else {
                                            const seqVal = mapViewMode === 'put' ? cell.putSeq : mapViewMode === 'check' ? cell.checkSeq : cell.pickSeq;
                                            const maxVal = mapViewMode === 'put' ? maxPut : mapViewMode === 'check' ? maxCheck : maxPick;
                                            bg = getHeatmapColor(seqVal, maxVal);
                                          }

                                          const isSelected = selectedLocationId === cell.id;

                                          return (
                                            <div
                                              key={cell.id}
                                              onClick={() => setSelectedLocationId(cell.id)}
                                              title={`${cell.id} | Bay ${cell.bay} Pos ${cell.position} Lvl ${cell.level} | Pick #${cell.pickSeq} Put #${cell.putSeq} Chk #${cell.checkSeq} | ${cell.isLocked ? 'LOCKED' : 'UNLOCKED'}`}
                                              style={{ background: bg }}
                                              className={`w-3 h-3 rounded-xs cursor-pointer transition hover:scale-125 hover:z-20 ${cell.isLocked ? 'ring-1 ring-amber-400' : ''} ${isSelected ? 'ring-2 ring-white scale-125 z-30' : ''} ${matchesSearch ? 'ring-2 ring-fuchsia-400' : ''}`}
                                            />
                                          );
                                        })}
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </div>

            {/* Right Cell Inspector & Diagnostics */}
            <div className="bg-[#0f172a] border-l border-[#334155] p-4 overflow-y-auto flex flex-col gap-4">
              <div className="bg-[#1e293b] border border-[#334155] p-3.5 rounded-xl">
                <h3 className="text-xs uppercase font-bold text-violet-400 tracking-wider mb-2">Location Inspector</h3>
                {selectedCell ? (
                  <div className="flex flex-col gap-1.5 font-mono text-xs">
                    <div className="text-sm font-bold text-white pb-1.5 border-b border-[#334155]">{selectedCell.id}</div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Aisle</span><span className="font-bold text-white">{selectedCell.aisle}</span></div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Bay</span><span className="text-white">{selectedCell.bay}</span></div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Position</span><span className="text-white">{selectedCell.position}</span></div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Level</span><span className="text-white">{selectedCell.level}</span></div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Side</span><span className="text-white">{selectedCell.side}</span></div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Lock Status</span><span className={selectedCell.isLocked ? 'text-amber-400 font-bold' : 'text-slate-300'}>{selectedCell.lockStatus}</span></div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Pick Seq</span><span className="text-violet-400 font-bold">{selectedCell.pickSeq}</span></div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Putaway Seq</span><span className="text-fuchsia-400 font-bold">{selectedCell.putSeq}</span></div>
                    <div className="grid grid-cols-2 py-1 border-b border-[#1e293b]"><span className="text-slate-400">Check Seq</span><span className="text-purple-400 font-bold">{selectedCell.checkSeq}</span></div>
                    <div className="grid grid-cols-2 py-1"><span className="text-slate-400">Site ID</span><span className="text-white">{selectedCell.siteId}</span></div>
                  </div>
                ) : (
                  <p className="text-xs text-slate-400 font-mono leading-relaxed">
                    Click any location cell on the warehouse map to inspect its position, sequence IDs, and lock flags.
                  </p>
                )}
              </div>

              {/* Map Statistics */}
              <div className="bg-[#1e293b] border border-[#334155] p-3.5 rounded-xl flex flex-col gap-2 font-mono text-xs">
                <h3 className="text-xs uppercase font-bold text-fuchsia-400 tracking-wider">Map Statistics</h3>
                <div className="flex justify-between py-1 border-b border-[#1e293b]"><span className="text-slate-400">Total Locations:</span><span className="text-white font-bold">{locations.length.toLocaleString()}</span></div>
                <div className="flex justify-between py-1 border-b border-[#1e293b]"><span className="text-slate-400">Total Aisles:</span><span className="text-white">{uniqueAisles.length}</span></div>
                <div className="flex justify-between py-1 border-b border-[#1e293b]"><span className="text-slate-400">Max Bays:</span><span className="text-white">{maxBayInMap}</span></div>
                <div className="flex justify-between py-1 border-b border-[#1e293b]"><span className="text-slate-400">Locked Locations:</span><span className="text-amber-400 font-bold">{locations.filter(l => l.isLocked).length.toLocaleString()}</span></div>
                <div className="flex justify-between py-1"><span className="text-slate-400">Max Pick Seq:</span><span className="text-violet-400 font-bold">{maxPick}</span></div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: SEQUENCE RULES & CONFIGURATION */}
        {activeTab === 'sequence' && (
          <div className="p-6 bg-[#020617] h-full overflow-auto flex flex-col gap-6 max-w-4xl mx-auto">
            <div>
              <h2 className="text-lg font-bold text-white mb-1">Sequence Routing & Traversal Rules</h2>
              <p className="text-xs text-slate-400">Configure parameters for Pick, Putaway, and Scissors-Lift Stock Check sequences.</p>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div className="bg-[#0f172a] border border-[#334155] p-4 rounded-xl flex flex-col gap-3">
                <h3 className="text-xs font-bold uppercase text-violet-400 tracking-wider">Sequence Numbering</h3>
                <div className="grid grid-cols-2 gap-3 text-xs">
                  <div>
                    <label className="text-[10px] text-slate-400 font-semibold">Start Sequence</label>
                    <input type="number" value={seqStart} onChange={(e) => setSeqStart(parseInt(e.target.value, 10) || 0)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white outline-none" />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-400 font-semibold">Sequence Increment Step</label>
                    <input type="number" value={seqStep} onChange={(e) => setSeqStep(parseInt(e.target.value, 10) || 1)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white outline-none" min="1" />
                  </div>
                </div>
              </div>

              <div className="bg-[#0f172a] border border-[#334155] p-4 rounded-xl flex flex-col gap-3">
                <h3 className="text-xs font-bold uppercase text-fuchsia-400 tracking-wider">Corridor Traversal Routing</h3>
                <div className="grid grid-cols-2 gap-3 text-xs">
                  <div>
                    <label className="text-[10px] text-slate-400 font-semibold">Pair Transition</label>
                    <select value={pairTransition} onChange={(e) => setPairTransition(e.target.value as any)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white outline-none">
                      <option value="uturn">U-Turn Into Next Pair</option>
                      <option value="straight">Straight Same Direction</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-400 font-semibold">Corridor Direction</label>
                    <select value={corridorDirection} onChange={(e) => setCorridorDirection(e.target.value as any)} className="w-full mt-1 bg-black border border-[#334155] rounded p-1.5 font-mono text-white outline-none">
                      <option value="forward">Bay 1 $\rightarrow$ Last Bay</option>
                      <option value="reverse">Last Bay $\rightarrow$ Bay 1</option>
                    </select>
                  </div>
                </div>
              </div>
            </div>

            {/* Explanation of the 3 sequence streams */}
            <div className="bg-[#0f172a] border border-[#334155] p-5 rounded-xl flex flex-col gap-3 font-mono text-xs">
              <h3 className="text-xs font-bold text-white uppercase tracking-wider">Active Sequence Streams</h3>
              <div className="space-y-3 text-slate-300">
                <div className="p-3 bg-black/40 border border-[#334155] rounded-lg">
                  <div className="text-violet-400 font-bold mb-1">1. Pick Sequence (`PICK_SEQUENCE`)</div>
                  <p className="text-[11px] text-slate-400">
                    Traverses corridor pairs facing each other across bays (Aisle 1 Bay 1 $\rightarrow$ Aisle 2 Bay 1 $\rightarrow$ Bay 2...), then U-turns into the next pair from the back of the warehouse.
                  </p>
                </div>
                <div className="p-3 bg-black/40 border border-[#334155] rounded-lg">
                  <div className="text-fuchsia-400 font-bold mb-1">2. Putaway Sequence (`PUT_SEQUENCE`)</div>
                  <p className="text-[11px] text-slate-400">
                    Exact mathematical inverse of the picking sequence, avoiding operator collision and traffic bottlenecks between pickers and replenishment fork-trucks.
                  </p>
                </div>
                <div className="p-3 bg-black/40 border border-[#334155] rounded-lg">
                  <div className="text-purple-400 font-bold mb-1">3. Stock Check Sequence (`CHECK_SEQUENCE`)</div>
                  <p className="text-[11px] text-slate-400">
                    Scissors-lift operational route: travels Level A across all bays (1 $\rightarrow$ N) of an aisle, then elevates to Level B across all bays, through all levels before transitioning to the next aisle.
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB PASSES: CUSTOM TRUCK ROUTES */}
        {activeTab === 'passes' && (
          <div className="p-6 bg-[#020617] h-full overflow-auto flex flex-col gap-4">
            <div>
              <h2 className="text-lg font-bold text-white mb-1">Pass Planner (Truck Routes)</h2>
              <p className="text-xs text-slate-400">Define the exact route passes the truck drives. Locations not captured by these passes will be pushed to the end.</p>
              {locations.length > 0 && locations[0].source === 'csv' && (
                <p className="text-xs text-fuchsia-400 font-bold mt-1">Unique Aisles in uploaded data: {Array.from(new Set(locations.map(l => l.aisle))).sort().join(', ')}</p>
              )}
            </div>
            
            <div className="border border-[#334155] rounded-xl overflow-hidden bg-[#1e293b] flex-1">
              <table className="w-full border-collapse font-mono text-xs text-left">
                <thead>
                  <tr className="bg-[#0f172a] text-slate-400 border-b border-[#334155]">
                    <th className="p-3">On</th>
                    <th className="p-3">Left Rack</th>
                    <th className="p-3">Right Rack</th>
                    <th className="p-3">From Bay</th>
                    <th className="p-3">To Bay</th>
                    <th className="p-3">Levels</th>
                    <th className="p-3">Bay Filter</th>
                    <th className="p-3">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#1e293b]">
                  {passes.map((pass, i) => (
                    <tr key={pass.id} className="hover:bg-white/5 transition">
                      <td className="p-3">
                        <input type="checkbox" checked={pass.on} onChange={e => { const np = [...passes]; np[i].on = e.target.checked; setPasses(np); }} className="w-4 h-4 rounded accent-violet-500" />
                      </td>
                      <td className="p-3">
                        <input type="text" value={pass.leftRack} onChange={e => { const np = [...passes]; np[i].leftRack = e.target.value.toUpperCase(); setPasses(np); }} className="w-20 bg-black border border-[#334155] rounded p-1 text-white outline-none" />
                      </td>
                      <td className="p-3">
                        <input type="text" value={pass.rightRack} onChange={e => { const np = [...passes]; np[i].rightRack = e.target.value.toUpperCase(); setPasses(np); }} className="w-20 bg-black border border-[#334155] rounded p-1 text-white outline-none" />
                      </td>
                      <td className="p-3">
                        <input type="number" value={pass.fromBay} onChange={e => { const np = [...passes]; np[i].fromBay = parseInt(e.target.value)||1; setPasses(np); }} className="w-16 bg-black border border-[#334155] rounded p-1 text-white outline-none" />
                      </td>
                      <td className="p-3">
                        <input type="number" value={pass.toBay} onChange={e => { const np = [...passes]; np[i].toBay = parseInt(e.target.value)||1; setPasses(np); }} className="w-16 bg-black border border-[#334155] rounded p-1 text-white outline-none" />
                      </td>
                      <td className="p-3">
                        <select value={pass.levelDirection} onChange={e => { const np = [...passes]; np[i].levelDirection = e.target.value as any; setPasses(np); }} className="bg-black border border-[#334155] rounded p-1 text-white outline-none">
                          <option value="up">A to Top</option>
                          <option value="down">Top to A</option>
                        </select>
                      </td>
                      <td className="p-3">
                        <select value={pass.bayFilter} onChange={e => { const np = [...passes]; np[i].bayFilter = e.target.value as any; setPasses(np); }} className="bg-black border border-[#334155] rounded p-1 text-white outline-none">
                          <option value="all">All</option>
                          <option value="even">Even Only</option>
                          <option value="odd">Odd Only</option>
                        </select>
                      </td>
                      <td className="p-3">
                        <div className="flex gap-2">
                          <button onClick={() => { if(i>0) { const np=[...passes]; [np[i-1], np[i]] = [np[i], np[i-1]]; setPasses(np); } }} className="text-slate-400 hover:text-white">↑</button>
                          <button onClick={() => { if(i<passes.length-1) { const np=[...passes]; [np[i+1], np[i]] = [np[i], np[i+1]]; setPasses(np); } }} className="text-slate-400 hover:text-white">↓</button>
                          <button onClick={() => setPasses(passes.filter((_, idx) => idx !== i))} className="text-rose-400 hover:text-rose-300"><Trash2 className="w-4 h-4"/></button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              
              <div className="p-4 border-t border-[#334155] flex justify-between">
                <button onClick={() => setPasses([...passes, { id: Date.now().toString(), on: true, leftRack: '', rightRack: '', fromBay: 1, toBay: 64, levelDirection: 'up', bayFilter: 'all' }])} className="px-3 py-1.5 bg-[#334155] hover:bg-slate-700 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 transition">
                  <Plus className="w-3.5 h-3.5" /> Add Pass
                </button>
                <button onClick={applyPassesToLocations} className="px-4 py-2 bg-amber-500 hover:bg-amber-400 text-black text-xs font-bold rounded-lg shadow-lg shadow-amber-500/20 transition">
                  Apply Passes & Sequence Data
                </button>
              </div>
            </div>
          </div>
        )}

        {/* TAB 4: DTO EXPORT & DATA TABLE */}
        {activeTab === 'export' && (
          <div className="h-full flex flex-col p-6 bg-[#020617] overflow-hidden gap-4">
            <div className="flex justify-between items-center shrink-0">
              <div>
                <h2 className="text-sm font-bold text-white">RedPrairie DLx Dispatcher 2016.2 — Data Take-On (DTO) Export</h2>
                <p className="text-xs text-slate-400">Export formatted records for Oracle `INTERFACE_LOCATION` database ingestion.</p>
              </div>
              <div className="flex gap-2">
                <button onClick={exportToExcel} className="px-3.5 py-1.5 bg-violet-500 hover:bg-violet-400 text-black font-extrabold rounded-lg text-xs transition flex items-center gap-1.5">
                  <Download className="w-3.5 h-3.5" /> Download .xlsx
                </button>
                <button onClick={exportToCsv} className="px-3.5 py-1.5 bg-[#334155] hover:bg-slate-700 text-white font-bold rounded-lg text-xs transition flex items-center gap-1.5">
                  <FileText className="w-3.5 h-3.5" /> Download .csv
                </button>
              </div>
            </div>

            {/* Records Data Table Preview */}
            <div className="flex-1 overflow-auto border border-[#334155] rounded-xl bg-[#0f172a]">
              <table className="w-full border-collapse font-mono text-xs">
                <thead>
                  <tr className="bg-[#0f172a] text-slate-400 border-b border-[#334155] sticky top-0 text-left">
                    <th className="p-2.5">PICK_SEQ</th>
                    <th className="p-2.5">PUT_SEQ</th>
                    <th className="p-2.5">CHECK_SEQ</th>
                    <th className="p-2.5">LOCATION_ID</th>
                    <th className="p-2.5">Aisle</th>
                    <th className="p-2.5">Bay</th>
                    <th className="p-2.5">Pos</th>
                    <th className="p-2.5">Level</th>
                    <th className="p-2.5">Side</th>
                    <th className="p-2.5">Locked?</th>
                    <th className="p-2.5">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#1e293b]">
                  {locations.length === 0 ? (
                    <tr>
                      <td colSpan={11} className="text-center text-slate-500 p-8">
                        No locations generated yet. Go to "Aisle Builder" and click "Build Map & Calculate Sequences".
                      </td>
                    </tr>
                  ) : (
                    locations.slice(0, 200).map((row, idx) => (
                      <tr key={idx} className={`hover:bg-white/5 ${row.isLocked ? 'bg-amber-500/10' : ''}`}>
                        <td className="p-2.5 text-violet-400 font-bold">{row.pickSeq}</td>
                        <td className="p-2.5 text-fuchsia-400">{row.putSeq}</td>
                        <td className="p-2.5 text-purple-400">{row.checkSeq}</td>
                        <td className="p-2.5 text-white font-bold">{row.id}</td>
                        <td className="p-2.5 text-slate-300">{row.aisle}</td>
                        <td className="p-2.5 text-slate-300">{row.bay}</td>
                        <td className="p-2.5 text-slate-300">{row.position}</td>
                        <td className="p-2.5 text-slate-300">{row.level}</td>
                        <td className="p-2.5 text-slate-400">{row.side}</td>
                        <td className="p-2.5">{row.isLocked ? <span className="text-amber-400 font-bold">Yes</span> : 'No'}</td>
                        <td className="p-2.5 text-slate-400">{row.lockStatus}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
