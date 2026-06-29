import { useCallback, useState, useEffect, useMemo, useRef } from 'react';
import AsyncStorage from '@react-native-async-storage/async-storage';
import {
  Alert,
  Image,
  Modal,
  Pressable,
  SafeAreaView,
  ScrollView,
  StyleSheet,
  Text,
  View,
  TextInput,
  KeyboardAvoidingView,
  Platform,
  FlatList,
  Dimensions,
  TouchableOpacity,
} from 'react-native';
import { StatusBar } from 'expo-status-bar';
import { LinearGradient } from 'expo-linear-gradient';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';
import * as XLSX from 'xlsx';
import DateTimePicker from '@react-native-community/datetimepicker';

import * as DocumentPicker from 'expo-document-picker';
import {
  parseExcelBase64,
  ExcelParseResult,
  PatientRecord,
  groupFieldsByCategory,
  DISPLAY_ORDER,
  CATEGORY_LABELS,
  OB_GYNE_SUB_CATEGORIES,
  FieldCategory,
} from './src/utils/excel';

const teamSilayanLogo = require('./assets/team-silayan.png');
const { width: screenWidth, height: screenHeight } = Dimensions.get('window');

// Import carousel images
const carouselImages = [
  require('./assets/1.png'),
  require('./assets/2.png'),
  require('./assets/3.png'),
  require('./assets/4.png'),
  require('./assets/5.png'),
  require('./assets/6.png'),
  require('./assets/7.png'),
];

type LoadedFile = {
  name: string;
  result: ExcelParseResult;
};

type PatientGroup = {
  name: string;
  records: PatientRecord[];
};

type DatePickerState = {
  visible: boolean;
  mode: 'single' | 'edit';
  targetId: string | null;
  currentValue: string;
  dateType: 'screening' | 'referral';
};

const STORAGE_KEY = '@team_silayan_patient_records';

// Helper function to categorize fields for display (new grouped structure)
function categorizeFieldsForDisplay(fields: Record<string, string>) {
  const screeningDates: string[] = [];
  const referralDates: string[] = [];
  let admitNumber = '';

  Object.entries(fields).forEach(([key, value]) => {
    const lower = key.toLowerCase().trim();
    if (lower.includes('screening date')) screeningDates.push(value);
    else if (lower.includes('referral date') || lower === 'date') referralDates.push(value);
    else if (lower.includes('admit number') || lower.includes('admit no')) admitNumber = value;
  });

  const grouped = groupFieldsByCategory(fields);

  return { grouped, screeningDates, referralDates, admitNumber };
}

// Helper to determine if patient is UR (both breasts UR)
function isPatientUr(rightBreast: string, leftBreast: string): boolean {
  return rightBreast?.toUpperCase() === 'UR' && leftBreast?.toUpperCase() === 'UR';
}

// ALL field keys from the Excel template
const EXCEL_FIELD_KEYS = [
  'ID no. (Hosp)',
  'Age',
  'Address',
  'Contact No.',
  'Civil Status',
  'No. of Children',
  'Menarche',
  'LMP',
  'AOG',
  'Menstrual Bleeding Pattern',
  'No. of pads/day',
  'Pregnancy History',
  'Age at first full term pregnancy',
  'Oral Contraceptives Use',
  'Duration of usage of oral contraceptives',
  'History of previous cervical cancer screening',
  'History of abnormal vaginal discharge',
  'History of abnormal vaginal bleeding',
  'Age of first intercourse',
  'No. of sexual partner',
  'Spouse/Partner/s',
  'Family History of Cancer',
  'Smoking',
  'Current medication',
  'Allergies',
  'Abdominal surgery',
  'BP',
  'Temp',
  'HR',
  'RR',
  'Height (cm)',
  'Weight (kg)',
  'BMI',
  'Skin',
  'HEENT',
  'Chest and Lungs',
  'Heart',
  'Abdomen',
];

// ─── Helper function to display field labels ──────────────────────────────────
function getDisplayLabel(key: string): string {
  const lower = key.toLowerCase().trim();
  if (lower === 'pregnancy history') return 'GTPAL';
  // Add override for ID No. (Hosp)
  if (lower === 'id no. (hosp)' || lower === 'id no. (hosp)') return 'ID No.';
  // Add other overrides if needed
  return key.replace(/\b\w/g, (c) => c.toUpperCase());
}

function App() {
  const [loadedFile, setLoadedFile] = useState<LoadedFile | null>(null);
  const [selectedPatientName, setSelectedPatientName] = useState<string | null>(null);
  const [aboutVisible, setAboutVisible] = useState(false);
  const [addPatientVisible, setAddPatientVisible] = useState(false);
  const [editRecordPatient, setEditRecordPatient] = useState<PatientRecord | null>(null);
  const [viewRecordPatient, setViewRecordPatient] = useState<PatientRecord | null>(null);
  const [addRecordForPatient, setAddRecordForPatient] = useState<string | null>(null);
  const [sidebarVisible, setSidebarVisible] = useState(false);
  const [staffsVisible, setStaffsVisible] = useState(false);
  const [statusBoardVisible, setStatusBoardVisible] = useState(false);
  const [importVisible, setImportVisible] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [searchScreeningDate, setSearchScreeningDate] = useState('');
  const [searchReferralDate, setSearchReferralDate] = useState('');
  const [carouselVisible, setCarouselVisible] = useState(true);
  const [currentSlide, setCurrentSlide] = useState(0);
  const [fullscreenVisible, setFullscreenVisible] = useState(false);
  const [fullscreenImageIndex, setFullscreenImageIndex] = useState(0);
  const flatListRef = useRef<FlatList>(null);
  const fullscreenFlatListRef = useRef<FlatList>(null);
  const [patientListPage, setPatientListPage] = useState(0);
  const PATIENTS_PER_PAGE = 20;
  const [patientDetailVisible, setPatientDetailVisible] = useState(false);
  const [detailPatientName, setDetailPatientName] = useState<string | null>(null);
  const [detailPatientRecords, setDetailPatientRecords] = useState<PatientRecord[]>([]);
  const [statusBoardPatientName, setStatusBoardPatientName] = useState<string | null>(null);

  // Load saved data on mount
  useEffect(() => {
    const loadData = async () => {
      try {
        const saved = await AsyncStorage.getItem(STORAGE_KEY);
        if (saved) {
          const parsed = JSON.parse(saved);
          if (parsed.name && parsed.result && parsed.result.patients) {
            // Backward compat: ensure screeningDate exists
            parsed.result.patients.forEach((p: any) => {
              if (!p.screeningDate) {
                p.screeningDate = p.date || '(no date)';
              }
              if (!p.rightBreastValue) {
                p.rightBreastValue = '(empty)';
              }
              if (!p.leftBreastValue) {
                p.leftBreastValue = '(empty)';
              }
            });
            setLoadedFile(parsed);
          }
        }
      } catch (e) {
        console.error('Failed to load data:', e);
      } finally {
        setIsLoading(false);
      }
    };
    loadData();
  }, []);

  // Auto-save whenever loadedFile changes
  useEffect(() => {
    const saveData = async () => {
      if (loadedFile) {
        try {
          await AsyncStorage.setItem(STORAGE_KEY, JSON.stringify(loadedFile));
        } catch (e) {
          console.error('Failed to save data:', e);
        }
      }
    };
    if (!isLoading) {
      saveData();
    }
  }, [loadedFile, isLoading]);

  useEffect(() => {
    if (patientDetailVisible && detailPatientName && loadedFile) {
      const refreshed = loadedFile.result.patients
        .filter(
          (p) =>
            p.patientName.toLowerCase().trim() ===
            detailPatientName.toLowerCase().trim(),
        )
        .sort((a, b) => {
          const dateA = new Date(a.date).getTime() || 0;
          const dateB = new Date(b.date).getTime() || 0;
          return dateA - dateB;
        });
      setDetailPatientRecords(refreshed);
    }
  }, [loadedFile, patientDetailVisible, detailPatientName]);

  // Auto-advance carousel
  useEffect(() => {
    if (!carouselVisible || fullscreenVisible) return;
    const timer = setInterval(() => {
      setCurrentSlide((prev) => {
        const next = (prev + 1) % carouselImages.length;
        flatListRef.current?.scrollToIndex({ index: next, animated: true });
        return next;
      });
    }, 3000);
    return () => clearInterval(timer);
  }, [carouselVisible, fullscreenVisible]);

  const patientGroups = useMemo<PatientGroup[]>(() => {
    if (!loadedFile?.result.patients.length) return [];
    const map = new Map<string, PatientRecord[]>();
    loadedFile.result.patients.forEach((p) => {
      const key = p.patientName.toLowerCase().trim();
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(p);
    });
    map.forEach((records) => {
      records.sort((a, b) => {
        const dateA = new Date(a.date).getTime() || 0;
        const dateB = new Date(b.date).getTime() || 0;
        return dateA - dateB;
      });
    });
    return Array.from(map.entries()).map(([name, records]) => ({
      name: records[0].patientName,
      records,
    }));
  }, [loadedFile]);

  const [sortMode, setSortMode] = useState<'name' | 'screeningDate' | 'referralDate' | 'status'>('name');
  const [sortAsc, setSortAsc] = useState(true);

  const parseSortDate = (dateStr: string): number => {
    if (!dateStr || dateStr === '(no date)' || dateStr === '(empty)') return 0;
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      const month = parseInt(parts[0], 10) - 1;
      const day = parseInt(parts[1], 10);
      const year = parseInt(parts[2], 10);
      if (!isNaN(month) && !isNaN(day) && !isNaN(year)) {
        return new Date(year, month, day).getTime();
      }
    }
    return new Date(dateStr).getTime() || 0;
  };

  const sortedPatientGroups = useMemo(() => {
    const sorted = [...patientGroups];
    sorted.sort((a, b) => {
      let cmp = 0;
      if (sortMode === 'name') {
        cmp = a.name.localeCompare(b.name);
      } else if (sortMode === 'screeningDate') {
        const latestA = a.records[a.records.length - 1];
        const latestB = b.records[b.records.length - 1];
        const dateA = parseSortDate(latestA?.screeningDate || '');
        const dateB = parseSortDate(latestB?.screeningDate || '');
        cmp = dateA - dateB;
      } else if (sortMode === 'referralDate') {
        const latestA = a.records[a.records.length - 1];
        const latestB = b.records[b.records.length - 1];
        const dateA = parseSortDate(latestA?.date || '');
        const dateB = parseSortDate(latestB?.date || '');
        cmp = dateA - dateB;
      } else if (sortMode === 'status') {
        const statusA = a.records[a.records.length - 1]?.isUr ? 1 : 0;
        const statusB = b.records[b.records.length - 1]?.isUr ? 1 : 0;
        cmp = statusA - statusB;
      }
      return sortAsc ? cmp : -cmp;
    });
    return sorted;
  }, [patientGroups, sortMode, sortAsc]);

  const filteredPatientGroups = useMemo(() => {
    let filtered = sortedPatientGroups;
    
    // Filter by name
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      filtered = filtered.filter((g) => g.name.toLowerCase().includes(q));
    }
    
    // Filter by screening date
    if (searchScreeningDate.trim()) {
      const q = searchScreeningDate.trim();
      filtered = filtered.filter((g) => {
        const latest = g.records[g.records.length - 1];
        const sd = latest.screeningDate || '';
        return sd.toLowerCase().includes(q.toLowerCase());
      });
    }
    
    // Filter by referral date
    if (searchReferralDate.trim()) {
      const q = searchReferralDate.trim();
      filtered = filtered.filter((g) => {
        const latest = g.records[g.records.length - 1];
        const rd = latest.date || '';
        return rd.toLowerCase().includes(q.toLowerCase());
      });
    }
    
    return filtered;
  }, [sortedPatientGroups, searchQuery, searchScreeningDate, searchReferralDate]);

  const selectedRecords = useMemo(() => {
    if (!selectedPatientName || !loadedFile) return [];
    return loadedFile.result.patients.filter(
      (p) => p.patientName.toLowerCase().trim() === selectedPatientName.toLowerCase().trim()
    ).sort((a, b) => {
      const dateA = new Date(a.date).getTime() || 0;
      const dateB = new Date(b.date).getTime() || 0;
      return dateA - dateB;
    });
  }, [selectedPatientName, loadedFile]);

  const handleSavePatient = useCallback(
    (patientData: {
      patientName: string;
      date: string;
      screeningDate: string;
      rightBreastValue: string;
      leftBreastValue: string;
      breastValue: string;
      fields: Record<string, string>;
    }) => {
      const isUr = isPatientUr(patientData.rightBreastValue, patientData.leftBreastValue);
      const maxRowIndex = loadedFile?.result.patients.reduce((max, p) => Math.max(max, p.rowIndex), 0) ?? 0;
      const nextRowIndex = maxRowIndex + 1;

      const newPatient: PatientRecord = {
        rowIndex: nextRowIndex,
        date: patientData.date,
        screeningDate: patientData.screeningDate,
        patientName: patientData.patientName,
        breastValue: patientData.breastValue || '(empty)',
        rightBreastValue: patientData.rightBreastValue || '(empty)',
        leftBreastValue: patientData.leftBreastValue || '(empty)',
        isUr,
        fields: patientData.fields,
      };

      if (loadedFile) {
        const updatedPatients = [...loadedFile.result.patients, newPatient];
        setLoadedFile({
          ...loadedFile,
          result: {
            ...loadedFile.result,
            patients: updatedPatients,
          },
        });
      } else {
        setLoadedFile({
          name: 'Patient Records',
          result: {
            sheetName: 'Patient Sheet',
            patients: [newPatient],
          },
        });
      }

      setSelectedPatientName(newPatient.patientName);
    },
    [loadedFile],
  );

  const handleUpdateRecord = useCallback(
    (updatedPatient: PatientRecord) => {
      if (!loadedFile) return;
      const updatedPatients = loadedFile.result.patients.map((p) =>
        p.rowIndex === updatedPatient.rowIndex ? updatedPatient : p,
      );
      setLoadedFile({
        ...loadedFile,
        result: { ...loadedFile.result, patients: updatedPatients },
      });
      setEditRecordPatient(null);
  
      // Refresh patient detail if it's open
      if (patientDetailVisible && detailPatientName) {
        // If the name changed, update detailPatientName to the new name
        const nameChanged =
          updatedPatient.patientName.toLowerCase().trim() !==
          detailPatientName.toLowerCase().trim();
        if (nameChanged) {
          setDetailPatientName(updatedPatient.patientName);
        }
  
        const searchName = nameChanged
          ? updatedPatient.patientName
          : detailPatientName;
  
        const refreshedRecords = updatedPatients
          .filter(
            (p) =>
              p.patientName.toLowerCase().trim() ===
              searchName.toLowerCase().trim(),
          )
          .sort((a, b) => {
            const dateA = new Date(a.date).getTime() || 0;
            const dateB = new Date(b.date).getTime() || 0;
            return dateA - dateB;
          });
        setDetailPatientRecords(refreshedRecords);
      }
    },
    [loadedFile, patientDetailVisible, detailPatientName],
  );

  const exportSingleRecord = useCallback(async (record: PatientRecord) => {
    try {
      const headers = Object.keys(record.fields);
      const row = headers.map((h) => record.fields[h] ?? '');
      const ws = XLSX.utils.aoa_to_sheet([headers, row]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Record');
      const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
      const uri = `${FileSystem.cacheDirectory}record_${record.rowIndex}_${Date.now()}.xlsx`;
      await FileSystem.writeAsStringAsync(uri, base64, { encoding: FileSystem.EncodingType.Base64 });
      const canShare = await Sharing.isAvailableAsync();
      if (canShare) {
        await Sharing.shareAsync(uri, {
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          dialogTitle: `Export Record - ${record.patientName}`,
        });
      } else {
        Alert.alert('Exported', `File saved to cache: ${uri}`);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Export failed.';
      Alert.alert('Export failed', msg);
    }
  }, []);

  const clearAllData = useCallback(() => {
    Alert.alert(
      'Clear All Data',
      'Are you sure you want to delete all patient records? This cannot be undone.',
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Clear All',
          style: 'destructive',
          onPress: async () => {
            try {
              await AsyncStorage.removeItem(STORAGE_KEY);
              setLoadedFile(null);
              setSelectedPatientName(null);
            } catch (e) {
              Alert.alert('Error', 'Failed to clear data.');
            }
          },
        },
      ]
    );
  }, []);

  const importExcel = useCallback(async (base64: string, fileName: string) => {
    try {
      const checkResult = parseExcelBase64(base64);
      // Inject admit numbers for patients that don't have one
      const admitNumberKey = Object.keys(checkResult.patients[0]?.fields ?? {}).find(
        (k) => k.toLowerCase().includes('admit number')
      );
      if (!admitNumberKey) {
        checkResult.patients.forEach((patient) => {
          patient.fields['Admit Number'] = String(Math.floor(100000000 + Math.random() * 900000000));
        });
      }
      // Ensure screeningDate is set and right/left breast values exist
      checkResult.patients.forEach((patient) => {
        if (!patient.screeningDate) {
          const sdKey = Object.keys(patient.fields).find(k => k.toLowerCase().includes('screening date'));
          patient.screeningDate = sdKey ? patient.fields[sdKey] : patient.date || '(no date)';
        }
        if (!patient.rightBreastValue) {
          const rbKey = Object.keys(patient.fields).find(k => k.toLowerCase().includes('right breast'));
          patient.rightBreastValue = rbKey ? patient.fields[rbKey] : '(empty)';
        }
        if (!patient.leftBreastValue) {
          const lbKey = Object.keys(patient.fields).find(k => k.toLowerCase().includes('left breast'));
          patient.leftBreastValue = lbKey ? patient.fields[lbKey] : '(empty)';
        }
        patient.isUr = isPatientUr(patient.rightBreastValue, patient.leftBreastValue);
      });
      if (loadedFile) {
        // Merge with existing data
        const existingRows = loadedFile.result.patients.map(p => p.rowIndex);
        const maxRow = Math.max(...existingRows, 0);
        const newPatients = checkResult.patients.map((p, i) => ({ ...p, rowIndex: maxRow + i + 1 }));
        setLoadedFile({
          ...loadedFile,
          result: {
            ...loadedFile.result,
            patients: [...loadedFile.result.patients, ...newPatients],
          },
        });
        Alert.alert('Imported', `Added ${newPatients.length} records from ${fileName}.`);
      } else {
        setLoadedFile({ name: fileName, result: checkResult });
        Alert.alert('Imported', `Loaded ${checkResult.patients.length} records from ${fileName}.`);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Import failed.';
      Alert.alert('Import failed', msg);
    }
  }, [loadedFile]);

  const exportExcel = useCallback(async () => {
    if (!loadedFile?.result.patients.length) {
      Alert.alert('Nothing to export', 'Create some patient records first.');
      return;
    }
    try {
      const patients = loadedFile.result.patients;
      // Get all unique field keys
      const allKeysSet = new Set<string>();
      patients.forEach((p) => Object.keys(p.fields).forEach((k) => allKeysSet.add(k)));
      // Put Screening Date and Referral Date first, then the rest
      const priorityKeys = ['Screening Date', 'Referral Date'];
      const otherKeys = Array.from(allKeysSet).filter(k => !priorityKeys.includes(k));
      const headers = [...priorityKeys.filter(k => allKeysSet.has(k)), ...otherKeys];

      // Build rows - one per patient record
      const rows: string[][] = [];
      patients.forEach((p) => {
        const row = headers.map((h) => p.fields[h] ?? '');
        rows.push(row);
      });

      const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, loadedFile.result.sheetName || 'Sheet1');
      const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
      const uri = `${FileSystem.cacheDirectory}export_${Date.now()}.xlsx`;
      await FileSystem.writeAsStringAsync(uri, base64, { encoding: FileSystem.EncodingType.Base64 });
      const canShare = await Sharing.isAvailableAsync();
      if (canShare) {
        await Sharing.shareAsync(uri, {
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          dialogTitle: 'Export Patient Data',
        });
      } else {
        Alert.alert('Exported', `File saved to cache: ${uri}`);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Export failed.';
      Alert.alert('Export failed', msg);
    }
  }, [loadedFile]);

  const { result } = loadedFile ?? {};

  const handleImagePress = (index: number) => {
    setFullscreenImageIndex(index);
    setFullscreenVisible(true);
  };

  const renderCarouselItem = ({ item, index }: { item: any; index: number }) => (
    <TouchableOpacity 
      style={styles.carouselSlide}
      onPress={() => handleImagePress(index)}
      activeOpacity={0.9}
    >
      <Image source={item} style={styles.carouselImage} resizeMode="contain" />
    </TouchableOpacity>
  );

  const renderFullscreenItem = ({ item, index }: { item: any; index: number }) => (
    <View style={styles.fullscreenSlide}>
      <Image source={item} style={styles.fullscreenImage} resizeMode="contain" />
    </View>
  );

  const renderDot = (index: number) => (
    <View
      key={index}
      style={[
        styles.carouselDot,
        currentSlide === index && styles.carouselDotActive,
      ]}
    />
  );

  const renderFullscreenDot = (index: number) => (
    <View
      key={index}
      style={[
        styles.fullscreenDot,
        fullscreenImageIndex === index && styles.fullscreenDotActive,
      ]}
    />
  );

  const todayString = () => {
    const today = new Date();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    const yyyy = today.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  };

  return (
    <SafeAreaView style={styles.safeArea}>
      <StatusBar style="light" />
      <View style={styles.container}>
        <Image
          source={require('./assets/team-silayan.png')}
          style={styles.backgroundLogo}
          resizeMode="contain"
        />
        <View style={styles.header}>
          <LinearGradient
            colors={['#CCCCCC', '#b51f55']}
            start={{ x: 0, y: 0 }}
            end={{ x: 1, y: 1 }}
            style={StyleSheet.absoluteFillObject}
          />
          <View style={styles.headerTop}>
            <View style={styles.headerLogoRow}>
              <Image
                source={require('./assets/team-silayan.png')}
                style={styles.headerIcon}
                resizeMode="contain"
              />
              <Image
                source={require('./assets/team-silayan-header.png')}
                style={styles.headerLogo}
                resizeMode="contain"
              />
            </View>
            <Pressable style={styles.burgerButton} onPress={() => setSidebarVisible(true)}>
              <View style={styles.burgerLine} />
              <View style={styles.burgerLine} />
              <View style={styles.burgerLine} />
            </Pressable>
          </View>
        </View>

        <View style={styles.buttonRow}>
          <Pressable
            style={[styles.button, styles.primaryButton, styles.rowButton]}
            onPress={() => setAddPatientVisible(true)}
          >
            <LinearGradient
              colors={['#db4278', '#b51f55']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 0 }}
              style={StyleSheet.absoluteFillObject}
            />
            <Text style={styles.primaryButtonText}>Add Patient</Text>
          </Pressable>

          <Pressable
            style={[styles.button, styles.exportButton, styles.rowButton]}
            onPress={exportExcel}
          >
            <Text style={styles.exportButtonText}>Export All Records</Text>
          </Pressable>
        </View>

        {isLoading ? (
          <View style={styles.placeholderCard}>
            <Text style={styles.placeholderTitle}>Loading...</Text>
            <Text style={styles.placeholderText}>Restoring your patient records...</Text>
          </View>
        ) : loadedFile && result ? (
          <ScrollView style={styles.resultScroll} contentContainerStyle={styles.resultSection}>
            <View style={styles.fileCard}>
              <Text style={styles.fileName}>{loadedFile.name}</Text>
              <Text style={styles.fileMeta}>
                {patientGroups.length} patient{patientGroups.length === 1 ? '' : 's'} ·{' '}
                {result.patients.length} record{result.patients.length === 1 ? '' : 's'} ·{' '}
                {result.sheetName}
              </Text>
            </View>

            {/* Search Section */}
            <View style={styles.searchSection}>
              <View style={styles.searchRow}>
                <TextInput
                  style={styles.searchInput}
                  placeholder="Search patient name..."
                  placeholderTextColor="#9ca3af"
                  value={searchQuery}
                  onChangeText={setSearchQuery}
                />
                {searchQuery.length > 0 && (
                  <Pressable style={styles.searchClearBtn} onPress={() => setSearchQuery('')}>
                    <Text style={styles.searchClearBtnText}>✕</Text>
                  </Pressable>
                )}
              </View>
              
              <View style={styles.dateSearchRow}>
                <View style={styles.dateSearchGroup}>
                  <Text style={styles.dateSearchLabel}>Screening:</Text>
                  <TextInput
                    style={styles.dateSearchInput}
                    placeholder="MM/DD/YYYY"
                    placeholderTextColor="#9ca3af"
                    value={searchScreeningDate}
                    onChangeText={setSearchScreeningDate}
                  />
                  {searchScreeningDate.length > 0 && (
                    <Pressable style={styles.searchClearBtnSmall} onPress={() => setSearchScreeningDate('')}>
                      <Text style={styles.searchClearBtnText}>✕</Text>
                    </Pressable>
                  )}
                </View>
                
                <View style={styles.dateSearchGroup}>
                  <Text style={styles.dateSearchLabel}>Referral:</Text>
                  <TextInput
                    style={styles.dateSearchInput}
                    placeholder="MM/DD/YYYY"
                    placeholderTextColor="#9ca3af"
                    value={searchReferralDate}
                    onChangeText={setSearchReferralDate}
                  />
                  {searchReferralDate.length > 0 && (
                    <Pressable style={styles.searchClearBtnSmall} onPress={() => setSearchReferralDate('')}>
                      <Text style={styles.searchClearBtnText}>✕</Text>
                    </Pressable>
                  )}
                </View>
              </View>

              <View style={styles.quickDateFilterRow}>
                <Pressable 
                  style={[styles.quickDateFilterBtn, !searchScreeningDate && !searchReferralDate && styles.quickDateFilterBtnActive]}
                  onPress={() => {
                    setSearchScreeningDate('');
                    setSearchReferralDate('');
                  }}
                >
                  <Text style={[styles.quickDateFilterBtnText, !searchScreeningDate && !searchReferralDate && styles.quickDateFilterBtnTextActive]}>All</Text>
                </Pressable>
                <Pressable 
                  style={[styles.quickDateFilterBtn, searchScreeningDate && styles.quickDateFilterBtnActive]}
                  onPress={() => {
                    setSearchScreeningDate(todayString());
                    setSearchReferralDate('');
                  }}
                >
                  <Text style={[styles.quickDateFilterBtnText, searchScreeningDate && styles.quickDateFilterBtnTextActive]}>Today's Screening</Text>
                </Pressable>
                <Pressable 
                  style={[styles.quickDateFilterBtn, searchReferralDate && styles.quickDateFilterBtnActive]}
                  onPress={() => {
                    setSearchReferralDate(todayString());
                    setSearchScreeningDate('');
                  }}
                >
                  <Text style={[styles.quickDateFilterBtnText, searchReferralDate && styles.quickDateFilterBtnTextActive]}>Today's Referral</Text>
                </Pressable>
              </View>
              
              {(searchScreeningDate || searchReferralDate) && (
                <View style={styles.searchFilterChip}>
                  <Text style={styles.searchFilterChipText}>
                    Filtering by date{searchScreeningDate ? ` · Screening: ${searchScreeningDate}` : ''}
                    {searchReferralDate ? ` · Referral: ${searchReferralDate}` : ''}
                  </Text>
                  <Pressable onPress={() => {
                    setSearchScreeningDate('');
                    setSearchReferralDate('');
                  }}>
                    <Text style={styles.searchFilterChipClear}>✕ Clear</Text>
                  </Pressable>
                </View>
              )}
            </View>

            <View style={styles.listCard}>
              <View style={styles.listHeader}>
                <Pressable style={[styles.listHeaderBtn, styles.listNameCol]} onPress={() => {
                  if (sortMode === 'name') setSortAsc(!sortAsc);
                  else { setSortMode('name'); setSortAsc(true); }
                }}>
                  <Text style={styles.listHeaderText}>NAME</Text>
                  {sortMode === 'name' && (
                    <View style={styles.sortBadge}>
                      <Text style={styles.sortBadgeText}>{sortAsc ? 'A-Z' : 'Z-A'}</Text>
                    </View>
                  )}
                </Pressable>
                <Pressable style={[styles.listHeaderBtn, styles.listStatusCol]} onPress={() => {
                  if (sortMode === 'status') setSortAsc(!sortAsc);
                  else { setSortMode('status'); setSortAsc(true); }
                }}>
                  <Text style={styles.listHeaderText}>STATUS</Text>
                  {sortMode === 'status' && (
                    <View style={styles.sortBadge}>
                      <Text style={styles.sortBadgeText}>{sortAsc ? 'WARN-UR' : 'UR-WARN'}</Text>
                    </View>
                  )}
                </Pressable>
                <Pressable style={[styles.listHeaderBtn, styles.listScreeningCol]} onPress={() => {
                  if (sortMode === 'screeningDate') setSortAsc(!sortAsc);
                  else { setSortMode('screeningDate'); setSortAsc(true); }
                }}>
                  <Text style={styles.listHeaderText}>SCREEN</Text>
                  {sortMode === 'screeningDate' && (
                    <View style={styles.sortBadge}>
                      <Text style={styles.sortBadgeText}>{sortAsc ? 'OLD' : 'NEW'}</Text>
                    </View>
                  )}
                </Pressable>
                <Pressable style={[styles.listHeaderBtn, styles.listReferralCol]} onPress={() => {
                  if (sortMode === 'referralDate') setSortAsc(!sortAsc);
                  else { setSortMode('referralDate'); setSortAsc(true); }
                }}>
                  <Text style={styles.listHeaderText}>REFER</Text>
                  {sortMode === 'referralDate' && (
                    <View style={styles.sortBadge}>
                      <Text style={styles.sortBadgeText}>{sortAsc ? 'OLD' : 'NEW'}</Text>
                    </View>
                  )}
                </Pressable>
              </View>

              <View style={styles.searchResultsInfo}>
                <Text style={styles.searchResultsInfoText}>
                  Showing {filteredPatientGroups.length} of {patientGroups.length} patients
                  {searchScreeningDate && ` (Screening: ${searchScreeningDate})`}
                  {searchReferralDate && ` (Referral: ${searchReferralDate})`}
                </Text>
              </View>

              {(() => {
                const start = patientListPage * PATIENTS_PER_PAGE;
                const end = start + PATIENTS_PER_PAGE;
                const pageGroups = filteredPatientGroups.slice(start, end);
                return pageGroups.map((group) => (
                  <PatientListRow
                    key={group.name.toLowerCase().trim()}
                    group={group}
                    selected={false}
                    onPress={() => {
                      const recs = loadedFile?.result.patients.filter(
                        (p) => p.patientName.toLowerCase().trim() === group.name.toLowerCase().trim()
                      ).sort((a, b) => {
                        const dateA = new Date(a.date).getTime() || 0;
                        const dateB = new Date(b.date).getTime() || 0;
                        return dateA - dateB;
                      }) || [];
                      setDetailPatientName(group.name);
                      setDetailPatientRecords(recs);
                      setPatientDetailVisible(true);
                    }}
                  />
                ));
              })()}
              {filteredPatientGroups.length > PATIENTS_PER_PAGE && (
                <View style={styles.paginationRow}>
                  <Pressable
                    style={[styles.paginationBtn, patientListPage === 0 && styles.paginationBtnDisabled]}
                    onPress={() => setPatientListPage((p) => Math.max(0, p - 1))}
                    disabled={patientListPage === 0}
                  >
                    <Text style={[styles.paginationBtnText, patientListPage === 0 && styles.paginationBtnTextDisabled]}>← Prev</Text>
                  </Pressable>
                  <Text style={styles.paginationText}>
                    Page {patientListPage + 1} of {Math.ceil(filteredPatientGroups.length / PATIENTS_PER_PAGE)}
                  </Text>
                  <Pressable
                    style={[styles.paginationBtn, (patientListPage + 1) * PATIENTS_PER_PAGE >= filteredPatientGroups.length && styles.paginationBtnDisabled]}
                    onPress={() => setPatientListPage((p) => p + 1)}
                    disabled={(patientListPage + 1) * PATIENTS_PER_PAGE >= filteredPatientGroups.length}
                  >
                    <Text style={[styles.paginationBtnText, (patientListPage + 1) * PATIENTS_PER_PAGE >= filteredPatientGroups.length && styles.paginationBtnTextDisabled]}>Next →</Text>
                  </Pressable>
                </View>
              )}
            </View>

            {/* Patient detail now opens in a modal */}
          </ScrollView>
        ) : (
          <View style={styles.placeholderCard}>
            <Text style={styles.placeholderTitle}>How it works</Text>
            <Text style={styles.placeholderText}>1. Tap "Add Patient" to create a new patient record</Text>
            <Text style={styles.placeholderText}>2. All patients appear in the list</Text>
            <Text style={styles.placeholderText}>3. List shows name, status, screening date, and referral date</Text>
            <Text style={styles.placeholderText}>4. Tap a name to see full admission history</Text>
            <Text style={styles.placeholderText}>5. Tap "View" to see individual record details</Text>
            <Text style={styles.placeholderText}>6. Tap "Edit" to modify any record data</Text>
          </View>
        )}
      </View>

      {/* Carousel as overlay on top of everything */}
      <Modal
        visible={carouselVisible}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setCarouselVisible(false)}
      >
        <View style={styles.carouselOverlay}>
          <View style={styles.carouselContainer}>
            <View style={styles.carouselHeader}>
              <Image
                source={require('./assets/team-silayan.png')}
                style={styles.carouselLogo}
                resizeMode="contain"
              />
              <Text style={styles.carouselTitle}>Welcome to Team Silayan</Text>
              <Text style={styles.carouselSubtitle}>Tap any image to view fullscreen</Text>
            </View>
            
            <FlatList
              ref={flatListRef}
              data={carouselImages}
              renderItem={renderCarouselItem}
              horizontal
              pagingEnabled
              showsHorizontalScrollIndicator={false}
              onMomentumScrollEnd={(event) => {
                const index = Math.round(event.nativeEvent.contentOffset.x / screenWidth);
                setCurrentSlide(index);
              }}
              keyExtractor={(item, index) => index.toString()}
              style={styles.carouselFlatList}
            />
            
            <View style={styles.carouselDots}>
              {carouselImages.map((_, index) => renderDot(index))}
            </View>
            
            <Pressable
              style={styles.carouselCloseBtn}
              onPress={() => setCarouselVisible(false)}
            >
              <LinearGradient
                colors={['#db4278', '#b51f55']}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 0 }}
                style={StyleSheet.absoluteFillObject}
              />
              <Text style={styles.carouselCloseBtnText}>Get Started</Text>
            </Pressable>
          </View>
        </View>
      </Modal>

      {/* Fullscreen Image Viewer */}
      <Modal
        visible={fullscreenVisible}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setFullscreenVisible(false)}
      >
        <View style={styles.fullscreenOverlay}>
          <View style={styles.fullscreenHeader}>
            <Text style={styles.fullscreenCounter}>
              {fullscreenImageIndex + 1} / {carouselImages.length}
            </Text>
            <Pressable
              style={styles.fullscreenCloseBtn}
              onPress={() => setFullscreenVisible(false)}
            >
              <Text style={styles.fullscreenCloseBtnText}>✕</Text>
            </Pressable>
          </View>
          
          <FlatList
            ref={fullscreenFlatListRef}
            data={carouselImages}
            renderItem={renderFullscreenItem}
            horizontal
            pagingEnabled
            showsHorizontalScrollIndicator={false}
            initialScrollIndex={fullscreenImageIndex}
            onMomentumScrollEnd={(event) => {
              const index = Math.round(event.nativeEvent.contentOffset.x / screenWidth);
              setFullscreenImageIndex(index);
            }}
            keyExtractor={(item, index) => index.toString()}
            style={styles.fullscreenFlatList}
            getItemLayout={(data, index) => ({
              length: screenWidth,
              offset: screenWidth * index,
              index,
            })}
          />
          
          <View style={styles.fullscreenDots}>
            {carouselImages.map((_, index) => renderFullscreenDot(index))}
          </View>
        </View>
      </Modal>

      <SidebarModal
        visible={sidebarVisible}
        onClose={() => setSidebarVisible(false)}
        onAbout={() => {
          setSidebarVisible(false);
          setAboutVisible(true);
        }}
        onStaffs={() => {
          setSidebarVisible(false);
          setStaffsVisible(true);
        }}
        onStatusBoard={() => {
          setSidebarVisible(false);
          setStatusBoardVisible(true);
        }}
        onImport={() => {
          setSidebarVisible(false);
          setImportVisible(true);
        }}
        onClearData={clearAllData}
        hasData={!!loadedFile}
      />
      <ImportExcelModal
        visible={importVisible}
        onClose={() => setImportVisible(false)}
        onImport={importExcel}
      />
      <AboutModal visible={aboutVisible} onClose={() => setAboutVisible(false)} />
      <StaffsModal visible={staffsVisible} onClose={() => setStaffsVisible(false)} />
      <StatusBoardModal
        visible={statusBoardVisible}
        onClose={() => {
          setStatusBoardVisible(false);
          setStatusBoardPatientName(null);
        }}
        patientGroups={patientGroups}
        onPatientPress={(patientName) => {
          setStatusBoardPatientName(patientName);
          
          const records = loadedFile?.result.patients.filter(
            (p) => p.patientName.toLowerCase().trim() === patientName.toLowerCase().trim()
          ).sort((a, b) => {
            const dateA = new Date(a.date).getTime() || 0;
            const dateB = new Date(b.date).getTime() || 0;
            return dateA - dateB;
          }) || [];
          
          setDetailPatientName(patientName);
          setDetailPatientRecords(records);
          setPatientDetailVisible(true);
        }}
      />
      <AddPatientModal
        visible={addPatientVisible}
        onClose={() => setAddPatientVisible(false)}
        onSave={(data) => {
          handleSavePatient(data);
          setAddPatientVisible(false);
        }}
        existingPatients={loadedFile?.result.patients}
      />
      {editRecordPatient && (
        <EditRecordModal
          patient={editRecordPatient}
          onClose={() => setEditRecordPatient(null)}
          onSave={handleUpdateRecord}
        />
      )}
      {viewRecordPatient && (
        <ViewRecordModal
          patient={viewRecordPatient}
          onClose={() => setViewRecordPatient(null)}
          onExport={() => exportSingleRecord(viewRecordPatient)}
        />
      )}
      {addRecordForPatient && (
        <AddRecordForPatientModal
          patientName={addRecordForPatient}
          onClose={() => {
            setAddRecordForPatient(null);
            // Reopen patient detail if it was hidden
            if (detailPatientName) setPatientDetailVisible(true);
          }}
          onSave={(data) => {
            handleSavePatient(data);
            setAddRecordForPatient(null);
            // Go straight back to the patient detail screen
            if (detailPatientName) {
              setPatientDetailVisible(true);
            }
          }}
          existingPatients={loadedFile?.result.patients}
        />
      )}
      {patientDetailVisible && detailPatientName && (
        <PatientDetailModal
          visible={patientDetailVisible}
          patientName={detailPatientName}
          records={detailPatientRecords}
          onClose={() => {
            setPatientDetailVisible(false);
            setDetailPatientName(null);
            setDetailPatientRecords([]);
            if (statusBoardPatientName) {
              setStatusBoardVisible(true);
              setStatusBoardPatientName(null);
            }
          }}
          onViewRecord={(p) => setViewRecordPatient(p)}
          onEditRecord={(p) => setEditRecordPatient(p)}
          onAddNewRecord={() => {
            setPatientDetailVisible(false);
            setAddRecordForPatient(detailPatientName);
          }}
          onDeleteRecord={(p) => {
            Alert.alert(
              'Delete Record',
              `Are you sure you want to delete this record for ${p.patientName}?\n\nAdmit No: ${p.fields['Admit Number'] || 'N/A'}`,
              [
                { text: 'Cancel', style: 'cancel' },
                {
                  text: 'Delete',
                  style: 'destructive',
                  onPress: () => {
                    if (!loadedFile) return;
                    const updatedPatients = loadedFile.result.patients.filter(
                      (patient) => patient.rowIndex !== p.rowIndex
                    );
                    setLoadedFile({
                      ...loadedFile,
                      result: { ...loadedFile.result, patients: updatedPatients },
                    });

                    const refreshedRecords = updatedPatients
                      .filter(
                        (pt) =>
                          pt.patientName.toLowerCase().trim() ===
                          p.patientName.toLowerCase().trim()
                      )
                      .sort((a, b) => {
                        const dateA = new Date(a.date).getTime() || 0;
                        const dateB = new Date(b.date).getTime() || 0;
                        return dateA - dateB;
                      });
                    setDetailPatientRecords(refreshedRecords);

                    if (refreshedRecords.length === 0) {
                      setPatientDetailVisible(false);
                      setDetailPatientName(null);
                      setDetailPatientRecords([]);

                      if (statusBoardPatientName) {
                        setStatusBoardVisible(true);
                        setStatusBoardPatientName(null);
                      }
                    }
                  },
                },
              ]
            );
          }}
        />
      )}
    </SafeAreaView>
  );
}

function PatientDetailModal({
  visible,
  patientName,
  records,
  onClose,
  onViewRecord,
  onEditRecord,
  onAddNewRecord,
  onDeleteRecord,
}: {
  visible: boolean;
  patientName: string;
  records: PatientRecord[];
  onClose: () => void;
  onViewRecord: (p: PatientRecord) => void;
  onEditRecord: (p: PatientRecord) => void;
  onAddNewRecord: () => void;
  onDeleteRecord: (p: PatientRecord) => void;
}) {
  return (
    <Modal visible={visible} animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <View style={styles.modalHeaderRow}>
          <Text style={styles.modalHeaderTitle}>Patient Details</Text>
          <Pressable style={styles.modalCancelButton} onPress={onClose}>
            <Text style={styles.modalCancelButtonText}>Close</Text>
          </Pressable>
        </View>
        <ScrollView style={styles.modalScrollView} contentContainerStyle={styles.modalScrollContent}>
          <PatientDetail
            patientName={patientName}
            records={records}
            onViewRecord={onViewRecord}
            onEditRecord={onEditRecord}
            onAddNewRecord={onAddNewRecord}
            onDeleteRecord={onDeleteRecord}
          />
        </ScrollView>
      </SafeAreaView>
    </Modal>
  );
}


function PatientListRow({
  group,
  selected,
  onPress,
}: {
  group: PatientGroup;
  selected: boolean;
  onPress: () => void;
}) {
  const latestRecord = group.records[group.records.length - 1];

  // Safety guard for empty records (e.g. after name change)
  if (!latestRecord) {
    return (
      <View style={{ padding: 20, alignItems: 'center' }}>
        <Text style={{ color: '#9ca3af', fontSize: 16 }}>
          Record not found. The patient name may have changed.
        </Text>
      </View>
    );
  }

  const screeningDate = latestRecord.screeningDate && latestRecord.screeningDate !== '(no date)' && latestRecord.screeningDate !== '(empty)'
    ? latestRecord.screeningDate.split(';')[0].trim()
    : '(no date)';
  
  const isUr = latestRecord.isUr;
  // Check if referral date exists (only matters for warnings)
  const hasReferral = latestRecord.date && 
                     latestRecord.date !== '(no date)' && 
                     latestRecord.date !== '(empty)';

  // Get the referral date to display
  const getReferralDate = () => {
    if (hasReferral) {
      return latestRecord.date.split(';')[0].trim();
    }
    return null;
  };

  const handleReferralPress = () => {
    if (!isUr && hasReferral) {
      const referralDate = getReferralDate();
      Alert.alert('Referral Date', `Referral Date: ${referralDate}`);
    }
  };

  return (
    <Pressable 
      style={[styles.listRow, selected && styles.listRowSelected]}
      onPress={onPress}
    >
      <View style={styles.listNameCol}>
        <Text style={[styles.listCell, styles.listNameText]} numberOfLines={1}>
          {group.name}
        </Text>
        {group.records.length > 1 && (
          <Text style={styles.listRecordCount}>{group.records.length} visits</Text>
        )}
      </View>
      <View style={[styles.listStatusCol, styles.statusBadgeWrap]}>
        <View style={[styles.statusBadge, isUr ? styles.statusBadgeOk : styles.statusBadgeWarn]}>
          <Text style={[styles.statusBadgeText, { color: isUr ? '#065f46' : '#92400e' }]}>
            {isUr ? '✓' : '⚠'}
          </Text>
        </View>
      </View>
      <Text style={[styles.listCell, styles.listScreeningCol]} numberOfLines={1}>
        {screeningDate}
      </Text>

      {/* Referral column: 
          - For UR: show nothing (no icon)
          - For warnings: show X if no referral, checkmark if has referral 
          - Clicking checkmark shows the referral date
      */}
      <View style={[styles.listReferralCol, { alignItems: 'center', justifyContent: 'center' }]}>
        {!isUr && (
          hasReferral ? (
            <Pressable onPress={handleReferralPress}>
              <Text style={{ color: '#10b981', fontSize: 14, fontWeight: '700' }}>✓</Text>
            </Pressable>
          ) : (
            <Text style={{ color: '#dc2626', fontSize: 14, fontWeight: '700' }}>✕</Text>
          )
        )}
      </View>
    </Pressable>
  );
}

function PatientDetail({
  patientName,
  records,
  onViewRecord,
  onEditRecord,
  onAddNewRecord,
  onDeleteRecord,
}: {
  patientName: string;
  records: PatientRecord[];
  onViewRecord: (p: PatientRecord) => void;
  onEditRecord: (p: PatientRecord) => void;
  onAddNewRecord: () => void;
  onDeleteRecord: (p: PatientRecord) => void;
}) {
  const latestRecord = records[records.length - 1];
  const latestCategorized = categorizeFieldsForDisplay(latestRecord.fields);
  const isUr = latestRecord.isUr;

  // Get display order with "Other Details" first
    // Get display order with "Other Details" first
    const displayOrder: FieldCategory[] = ['other' as FieldCategory, ...DISPLAY_ORDER.filter(cat => cat !== 'other')];

  return (
    <View style={styles.detailSection}>
      <View style={styles.detailHeaderStack}>
        <View style={styles.detailNameContainer}>
          <Text style={styles.detailSectionTitle} numberOfLines={1} ellipsizeMode="tail">{patientName}</Text>
          <Text style={styles.detailSectionMeta}>
            {records.length} admission{records.length === 1 ? '' : 's'}
          </Text>
        </View>
        <Pressable style={styles.addNewRecordBtn} onPress={onAddNewRecord}>
          <LinearGradient
            colors={['#db4278', '#b51f55']}
            start={{ x: 0, y: 0 }}
            end={{ x: 1, y: 0 }}
            style={StyleSheet.absoluteFillObject}
          />
          <Text style={styles.addNewRecordBtnText} numberOfLines={1}>+ New Record</Text>
        </Pressable>
      </View>

      <AdmissionHistoryCard
        records={records}
        onViewRecord={onViewRecord}
        onEditRecord={onEditRecord}
        onDeleteRecord={onDeleteRecord}
      />
      <View style={styles.detailCard}>
        <Text style={styles.detailCardTitle}>Latest Record Details</Text>

        {/* Breast Findings */}
        <Text style={styles.categorySubLabel}>Breast Findings</Text>
        <DetailRow label="Right Breast" value={latestRecord.rightBreastValue || '(empty)'} />
        <DetailRow label="Left Breast"  value={latestRecord.leftBreastValue  || '(empty)'} />

        {/* All other grouped sections - Other Details first */}
        {displayOrder
          .filter((cat) => cat !== 'breastRight' && cat !== 'breastLeft')
          .map((cat) => {
            const entries = Object.entries(latestCategorized.grouped[cat] || {});
            if (entries.length === 0) return null;

            return (
              <View key={cat}>
                <Text style={styles.categorySubLabel}>
                  {CATEGORY_LABELS[cat]}
                </Text>

                {entries
                .filter(([_, value]) => value && value !== '(empty)' && value !== '')
                .map(([label, value]) => (
                  <DetailRow key={label} label={label} value={value} />
                ))}
              </View>
            );
          })}
      </View>
    </View>
  );
}

function AdmissionHistoryCard({
  records,
  onViewRecord,
  onEditRecord,
  onDeleteRecord,
}: {
  records: PatientRecord[];
  onViewRecord: (p: PatientRecord) => void;
  onEditRecord: (p: PatientRecord) => void;
  onDeleteRecord: (p: PatientRecord) => void;
}) {
  const [historySortAsc, setHistorySortAsc] = useState(true);
  const [historySearch, setHistorySearch] = useState('');
  const [historyPage, setHistoryPage] = useState(0);
  const HISTORY_PER_PAGE = 10;

  useEffect(() => {
    setHistoryPage(0);
  }, [records]);

  const sortedRecords = useMemo(() => {
    const arr = [...records];
    arr.sort((a, b) => {
      const parse = (d: string) => {
        const parts = d?.split('/');
        if (parts?.length === 3) return new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1])).getTime();
        return new Date(d || '').getTime() || 0;
      };
      const da = parse(a.screeningDate || a.date || '');
      const db = parse(b.screeningDate || b.date || '');
      return historySortAsc ? da - db : db - da;
    });
    return arr;
  }, [records, historySortAsc]);

  const filteredRecords = useMemo(() => {
    if (!historySearch.trim()) return sortedRecords;
    const q = historySearch.toLowerCase().trim();
    return sortedRecords.filter((r) => {
      const sd = (r.screeningDate || '').toLowerCase();
      const rd = (r.date || '').toLowerCase();
      const admitNo = (r.fields['Admit Number'] || '').toLowerCase();
      return sd.includes(q) || rd.includes(q) || admitNo.includes(q);
    });
  }, [sortedRecords, historySearch]);

  return (
    <View style={styles.admissionHistoryCard}>
      <View style={styles.admissionHistoryCardHeader}>
        <Text style={styles.admissionHistoryTitle}>Admission History</Text>
        <Pressable
          style={styles.historySortBtn}
          onPress={() => setHistorySortAsc(!historySortAsc)}
        >
          <Text style={styles.historySortBtnText}>{historySortAsc ? 'Oldest First ↑' : 'Newest First ↓'}</Text>
        </Pressable>
      </View>
      <View style={styles.historySearchRow}>
        <TextInput
          style={styles.historySearchInput}
          placeholder="Search by date or admit no..."
          placeholderTextColor="#9ca3af"
          value={historySearch}
          onChangeText={setHistorySearch}
        />
        {historySearch.length > 0 && (
          <Pressable style={styles.searchClearBtn} onPress={() => setHistorySearch('')}>
            <Text style={styles.searchClearBtnText}>✕</Text>
          </Pressable>
        )}
      </View>
      {(() => {
        const start = historyPage * HISTORY_PER_PAGE;
        const end = start + HISTORY_PER_PAGE;
        const pageRecords = filteredRecords.slice(start, end);
        return pageRecords.map((record, idx) => {
          const categorized = categorizeFieldsForDisplay(record.fields);
          const recordIsUr = record.isUr;
          return (
            <View key={record.rowIndex} style={styles.admissionHistoryRow}>
              <Text style={styles.admissionHistoryNumber}>#{start + idx + 1}</Text>
              <View style={styles.admissionHistoryInfo}>
                <Text style={styles.admissionHistoryAdmitNo}>Admit No: {categorized.admitNumber || 'N/A'}</Text>
                {categorized.screeningDates.length > 0 && (
                  <Text style={styles.admissionHistoryDate}>
                    Screening: {categorized.screeningDates.join(', ')}
                  </Text>
                )}
                {!recordIsUr && categorized.referralDates.length > 0 && (
                  <Text style={styles.admissionHistoryDate}>
                    Referral: {categorized.referralDates.join(', ')}
                  </Text>
                )}
                {recordIsUr && (
                  <Text style={[styles.admissionHistoryDate, { color: '#065f46', fontStyle: 'italic' }]}>
                    (UR - no referral needed)
                  </Text>
                )}
                <Text style={styles.admissionHistoryStatus}>
                  Right: {record.rightBreastValue} | Left: {record.leftBreastValue}
                </Text>
                <View style={styles.admissionActionRow}>
                  <Pressable style={styles.viewRecordBtn} onPress={() => onViewRecord(record)}>
                    <Text style={styles.viewRecordBtnText}>View</Text>
                  </Pressable>
                  <Pressable style={styles.editRecordBtn} onPress={() => onEditRecord(record)}>
                    <Text style={styles.editRecordBtnText}>Edit</Text>
                  </Pressable>
                  <Pressable style={styles.deleteRecordBtn} onPress={() => onDeleteRecord(record)}>
                    <Text style={styles.deleteRecordBtnText}>Delete</Text>
                  </Pressable>
                </View>
              </View>
            </View>
          );
        });
      })()}
      {filteredRecords.length === 0 && (
        <Text style={{ color: '#9ca3af', fontSize: 14, fontStyle: 'italic', padding: 8 }}>No records match your search.</Text>
      )}
      {filteredRecords.length > HISTORY_PER_PAGE && (
        <View style={styles.paginationRow}>
          <Pressable
            style={[styles.paginationBtn, historyPage === 0 && styles.paginationBtnDisabled]}
            onPress={() => setHistoryPage((p) => Math.max(0, p - 1))}
            disabled={historyPage === 0}
          >
            <Text style={[styles.paginationBtnText, historyPage === 0 && styles.paginationBtnTextDisabled]}>← Prev</Text>
          </Pressable>
          <Text style={styles.paginationText}>
            Page {historyPage + 1} of {Math.ceil(filteredRecords.length / HISTORY_PER_PAGE)}
          </Text>
          <Pressable
            style={[styles.paginationBtn, (historyPage + 1) * HISTORY_PER_PAGE >= filteredRecords.length && styles.paginationBtnDisabled]}
            onPress={() => setHistoryPage((p) => p + 1)}
            disabled={(historyPage + 1) * HISTORY_PER_PAGE >= filteredRecords.length}
          >
            <Text style={[styles.paginationBtnText, (historyPage + 1) * HISTORY_PER_PAGE >= filteredRecords.length && styles.paginationBtnTextDisabled]}>Next →</Text>
          </Pressable>
        </View>
      )}
    </View>
  );
}

function DetailRow({
  label,
  value,
  highlight = false,
}: {
  label: string;
  value: string;
  highlight?: boolean;
}) {
  const isBreastField = label?.toLowerCase?.() === 'right breast' || 
                         label?.toLowerCase?.() === 'left breast';
  const isUr = value?.toUpperCase?.() === 'UR';
  
  return (
    <View style={styles.detailRow}>
      <Text style={styles.detailLabel}>{getDisplayLabel(label)}</Text>
      <Text 
        style={[
          styles.detailValue, 
          highlight && styles.detailValueHighlight,
          isBreastField && (isUr ? styles.detailValueUr : styles.detailValueWarning)
        ]}
      >
        {value}
      </Text>
    </View>
  );
}

function AboutModal({ visible, onClose }: { visible: boolean; onClose: () => void }) {
  return (
    <Modal visible={visible} animationType="fade" transparent onRequestClose={onClose}>
      <View style={styles.aboutOverlay}>
        <Pressable style={styles.aboutBackdrop} onPress={onClose} />
        <View style={styles.aboutCard}>
          <Image source={teamSilayanLogo} style={styles.aboutLogo} resizeMode="contain" />
          <Text style={styles.aboutDescription}>
          Silay is a community-centered offline digital platform for Brgy. Tagasilay that transforms breast and cervical cancer screening into care that sees every woman through.
          </Text>
          <Text style={styles.aboutCredit}>Created by: Team Silayan</Text>
          <Pressable style={styles.aboutCloseButton} onPress={onClose}>
            <LinearGradient
              colors={['#db4278', '#b51f55']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 0 }}
              style={StyleSheet.absoluteFillObject}
            />
            <Text style={styles.aboutCloseButtonText}>Close</Text>
          </Pressable>
        </View>
      </View>
    </Modal>
  );
}

function ViewRecordModal({
  patient,
  onClose,
  onExport,
}: {
  patient: PatientRecord;
  onClose: () => void;
  onExport: () => void;
}) {
  const categorized = categorizeFieldsForDisplay(patient.fields);
  const isUr = patient.isUr;
  const hasReferral = categorized.referralDates.length > 0 && 
                     categorized.referralDates[0] !== '(no date)' && 
                     categorized.referralDates[0] !== '(empty)';

  // Get display order with "Other Details" first
  const displayOrder: FieldCategory[] = ['other' as FieldCategory, ...DISPLAY_ORDER.filter(cat => cat !== 'other')];

  return (
    <Modal visible animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <View style={styles.modalHeaderRow}>
          <Text style={styles.modalHeaderTitle}>Record Details</Text>
          <Pressable style={styles.modalCancelButton} onPress={onClose}>
            <Text style={styles.modalCancelButtonText}>Close</Text>
          </Pressable>
        </View>

        <ScrollView style={styles.modalScrollView} contentContainerStyle={styles.modalScrollContent}>
          <View style={styles.viewRecordHeader}>
            <Text style={styles.viewRecordName}>{patient.patientName}</Text>
            <Text style={styles.viewRecordMeta}>Row {patient.rowIndex} · Admit No: {categorized.admitNumber || 'N/A'}</Text>
          </View>

          <View
            style={[
              styles.statusCard,
              isUr ? styles.statusCardOk : styles.statusCardWarn,
              { marginBottom: 16 },
            ]}
          >
            <LinearGradient
              colors={isUr ? ['#eafaf1', '#d1f2e1'] : ['#fdf6e2', '#f9e8be']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 1 }}
              style={StyleSheet.absoluteFillObject}
            />
            <Text style={[styles.statusIcon, { color: isUr ? '#10b981' : '#f59e0b', fontSize: 48 }]}>
              {isUr ? '✓' : '⚠'}
            </Text>
            <Text style={[styles.statusTitle, { color: isUr ? '#065f46' : '#92400e' }]}>
              {isUr ? 'Breast Status: UR' : 'Breast Status: Suspicious'}
            </Text>
            <Text style={[styles.statusMessage, { color: isUr ? '#0f5132' : '#664d03' }]}>
              Right: {patient.rightBreastValue} | Left: {patient.leftBreastValue}
            </Text>
            {isUr && (
              <Text style={[styles.statusMessage, { color: '#065f46', fontStyle: 'italic', fontSize: 13 }]}>
                (No referral date required)
              </Text>
            )}
          </View>

          {/* Only show referral dates if NOT UR */}
          {!isUr && (
            <View style={styles.viewRecordSection}>
              <Text style={styles.viewRecordSectionTitle}>Referral Dates</Text>
              {categorized.referralDates.length > 0 && 
               categorized.referralDates[0] !== '(no date)' && 
               categorized.referralDates[0] !== '(empty)' ? (
                categorized.referralDates.map((date, idx) => (
                  <Text key={idx} style={styles.viewRecordItem}>• {date}</Text>
                ))
              ) : (
                <Text style={[styles.viewRecordEmpty, { color: '#dc2626' }]}>✕ No referral date recorded</Text>
              )}
            </View>
          )}

          {/* Breast Findings */}
          <View style={styles.viewRecordSection}>
            <Text style={styles.viewRecordSectionTitle}>Breast Findings</Text>
            <View style={styles.viewRecordFieldRow}>
              <Text style={styles.viewRecordFieldLabel}>Right Breast</Text>
              <Text 
                style={[
                  styles.viewRecordFieldValue,
                  patient.rightBreastValue?.toUpperCase?.() === 'UR' 
                    ? styles.viewRecordFieldValueUr 
                    : styles.viewRecordFieldValueWarning
                ]}
              >
                {patient.rightBreastValue || '(empty)'}
              </Text>
            </View>
            <View style={styles.viewRecordFieldRow}>
              <Text style={styles.viewRecordFieldLabel}>Left Breast</Text>
              <Text 
                style={[
                  styles.viewRecordFieldValue,
                  patient.leftBreastValue?.toUpperCase?.() === 'UR' 
                    ? styles.viewRecordFieldValueUr 
                    : styles.viewRecordFieldValueWarning
                ]}
              >
                {patient.leftBreastValue || '(empty)'}
              </Text>
            </View>
          </View>

          {/* Filter OUT breast fields (they're already displayed above) */}
          {displayOrder.filter(
            (cat) => cat !== 'breastRight' && cat !== 'breastLeft'
          ).map((cat) => {
            const entries = Object.entries(categorized.grouped[cat] || {});
            if (entries.length === 0) return null;
            return (
              <View key={cat} style={styles.viewRecordSection}>
                <Text style={styles.viewRecordSectionTitle}>{CATEGORY_LABELS[cat]}</Text>
                {entries
                .filter(([_, value]) => value && value !== '(empty)' && value !== '')
                .map(([label, value]) => (
                  <View key={label} style={styles.viewRecordFieldRow}>
                    <Text style={styles.viewRecordFieldLabel}>{getDisplayLabel(label)}</Text>
                    <Text style={styles.viewRecordFieldValue}>{value || '(empty)'}</Text>
                  </View>
                ))}
              </View>
            );
          })}

          <Pressable style={styles.exportSingleBtn} onPress={onExport}>
            <LinearGradient
              colors={['#db4278', '#b51f55']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 0 }}
              style={StyleSheet.absoluteFillObject}
            />
            <Text style={styles.exportSingleBtnText}>Export This Record</Text>
          </Pressable>
        </ScrollView>
      </SafeAreaView>
    </Modal>
  );
}

function EditRecordModal({
  patient,
  onClose,
  onSave,
}: {
  patient: PatientRecord;
  onClose: () => void;
  onSave: (patient: PatientRecord) => void;
}) {
  // Guard against undefined patient
  if (!patient || !patient.fields) {
    return null;
  }
  const [name, setName] = useState(patient.patientName);
  const [screeningDates, setScreeningDates] = useState<Array<{ id: string; value: string }>>([]);
  const [referralDates, setReferralDates] = useState<Array<{ id: string; value: string }>>([]);
  const [rightBreast, setRightBreast] = useState(patient.rightBreastValue || '(empty)');
  const [leftBreast, setLeftBreast] = useState(patient.leftBreastValue || '(empty)');
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const [datePicker, setDatePicker] = useState<DatePickerState>({
    visible: false,
    mode: 'edit',
    targetId: null,
    currentValue: '',
    dateType: 'screening',
  });

  const isUr = isPatientUr(rightBreast, leftBreast);

  // Get field keys from the patient
  const fieldKeys = useMemo(() => {
    return Object.keys(patient.fields).filter((key) => {
      const lowerKey = key.toLowerCase();
      return (
        !lowerKey.includes('name') &&
        !lowerKey.includes('date') &&
        !lowerKey.includes('breast') &&
        !lowerKey.includes('admit number') &&
        !lowerKey.includes('admit no') &&
        !lowerKey.includes('screening')
      );
    });
  }, [patient]);

  useEffect(() => {
    setName(patient.patientName);
    setRightBreast(patient.rightBreastValue || '(empty)');
    setLeftBreast(patient.leftBreastValue || '(empty)');

    const rawScreening = patient.screeningDate || '';
    if (rawScreening && rawScreening !== '(no date)' && rawScreening !== '(empty)') {
      const dates = rawScreening.split(';').map((d) => d.trim()).filter(Boolean);
      setScreeningDates(dates.map((d, i) => ({ id: `screening-${i}`, value: d })));
    } else {
      setScreeningDates([{ id: 'screening-init', value: todayString() }]);
    }

    const rawReferral = patient.date || '';
    if (rawReferral && rawReferral !== '(no date)' && rawReferral !== '(empty)') {
      const dates = rawReferral.split(';').map((d) => d.trim()).filter(Boolean);
      setReferralDates(dates.map((d, i) => ({ id: `referral-${i}`, value: d })));
    } else {
      setReferralDates([]);
    }

    const otherFields: Record<string, string> = {};
    Object.entries(patient.fields).forEach(([key, value]) => {
      const lower = key.toLowerCase();
      if (!lower.includes('name') && !lower.includes('date') && !lower.includes('breast') &&
          !lower.includes('admit number') && !lower.includes('admit no') && !lower.includes('screening')) {
        otherFields[key] = value;
      }
    });
    setFieldValues(otherFields);
  }, [patient]);

  const addScreeningDate = () => {
    setScreeningDates((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), value: todayString() },
    ]);
  };

  const updateScreeningDate = (id: string, value: string) => {
    setScreeningDates((prev) => prev.map((d) => (d.id === id ? { ...d, value } : d)));
  };

  const removeScreeningDate = (id: string) => {
    setScreeningDates((prev) => prev.filter((d) => d.id !== id));
  };

  const addReferralDate = () => {
    setReferralDates((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), value: todayString() },
    ]);
  };

  const updateReferralDate = (id: string, value: string) => {
    setReferralDates((prev) => prev.map((d) => (d.id === id ? { ...d, value } : d)));
  };

  const removeReferralDate = (id: string) => {
    setReferralDates((prev) => prev.filter((d) => d.id !== id));
  };

  const openDatePicker = (id: string, currentValue: string, dateType: 'screening' | 'referral') => {
    setDatePicker({
      visible: true,
      mode: 'edit',
      targetId: id,
      currentValue,
      dateType,
    });
  };

  const onDateChange = (event: any, selectedDate?: Date) => {
    if (selectedDate) {
      const mm = String(selectedDate.getMonth() + 1).padStart(2, '0');
      const dd = String(selectedDate.getDate()).padStart(2, '0');
      const yyyy = selectedDate.getFullYear();
      const formatted = `${mm}/${dd}/${yyyy}`;

      if (datePicker.mode === 'edit' && datePicker.targetId) {
        if (datePicker.dateType === 'screening') {
          updateScreeningDate(datePicker.targetId, formatted);
        } else {
          updateReferralDate(datePicker.targetId, formatted);
        }
      }
    }
    setDatePicker((prev) => ({ ...prev, visible: false }));
  };

  const handleSave = () => {
    if (!name.trim()) {
      Alert.alert('Validation Error', 'Patient Name is required.');
      return;
    }

    const combinedScreening = screeningDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryScreening = screeningDates[0]?.value.trim() || '(no date)';

    const combinedReferral = referralDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryReferral = referralDates[0]?.value.trim() || '(no date)';

    const updatedFields: Record<string, string> = { ...fieldValues };

    const allKeys = Object.keys(patient.fields);
    const nameKey = allKeys.find((k) => k.toLowerCase() === 'name' || k.toLowerCase().includes('patient name')) || 'name';
    const dateKey = allKeys.find((k) => k.toLowerCase() === 'date' || k.toLowerCase().includes('referral date')) || 'date';
    const screeningKey = allKeys.find((k) => k.toLowerCase().includes('screening date')) || 'Screening Date';
    const rightBreastKey = allKeys.find((k) => k.toLowerCase().includes('right breast')) || 'Right Breast';
    const leftBreastKey = allKeys.find((k) => k.toLowerCase().includes('left breast')) || 'Left Breast';

    updatedFields[nameKey] = name.trim();
    updatedFields[dateKey] = combinedReferral;
    updatedFields[screeningKey] = combinedScreening;
    updatedFields[rightBreastKey] = rightBreast.trim();
    updatedFields[leftBreastKey] = leftBreast.trim();

    const admitNoKey = allKeys.find((k) => k.toLowerCase().includes('admit number'));
    if (admitNoKey) {
      updatedFields[admitNoKey] = patient.fields[admitNoKey];
    }

    // Compute breastValue for display
    const rightIsUr = rightBreast.toUpperCase() === 'UR';
    const leftIsUr = leftBreast.toUpperCase() === 'UR';
    const isUrComputed = rightIsUr && leftIsUr;
    const breastValue = isUrComputed
      ? 'UR'
      : [
          !rightIsUr && rightBreast ? `R: ${rightBreast}` : '',
          !leftIsUr && leftBreast ? `L: ${leftBreast}` : '',
        ]
          .filter(Boolean)
          .join(' | ') || '(empty)';

    const updatedPatient: PatientRecord = {
      ...patient,
      patientName: name.trim(),
      date: primaryReferral,
      screeningDate: primaryScreening,
      breastValue,
      rightBreastValue: rightBreast.trim(),
      leftBreastValue: leftBreast.trim(),
      isUr: isUrComputed,
      fields: updatedFields,
    };

    onSave(updatedPatient);
  };

  // Categorize fields for grouped display
  const {
    menstrualHistory: editMenstrualHistory,
    pregnancyHistory: editPregnancyHistory,
    contraceptiveUse: editContraceptiveUse,
    sexualHistory: editSexualHistory,
    familySocialHistory: editFamilySocialHistory,
    medicalHistory: editMedicalHistory,
    vitalSigns: editVitalSigns,
    anthropometric: editAnthropometric,
    physicalExam: editPhysicalExam,
    demographics: editDemographics,
    other: editOtherKeys,
  } = categorizeFieldKeys(fieldKeys);

  return (
    <Modal visible animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <KeyboardAvoidingView
          behavior={Platform.OS === 'ios' ? 'padding' : 'height'}
          style={{ flex: 1 }}
        >
          <View style={styles.modalHeaderRow}>
            <Text style={styles.modalHeaderTitle}>Edit Record</Text>
            <Pressable style={styles.modalCancelButton} onPress={onClose}>
              <Text style={styles.modalCancelButtonText}>Cancel</Text>
            </Pressable>
          </View>

          <ScrollView style={styles.modalScrollView} contentContainerStyle={styles.modalScrollContent}>
            <Text style={styles.editRecordPatientName}>{patient.patientName}</Text>
            <Text style={styles.editRecordPatientMeta}>Row {patient.rowIndex} · Admit No: {patient.fields['Admit Number'] || 'N/A'}</Text>

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Patient Name *</Text>
              <TextInput
                style={styles.formInput}
                placeholder="Enter patient name"
                placeholderTextColor="#64748b"
                value={name}
                onChangeText={setName}
              />
            </View>

            <View style={styles.sectionHeaderRow}>
              <Text style={styles.sectionHeaderTitle}>Screening Dates *</Text>
              <Pressable style={styles.addDateBtn} onPress={addScreeningDate}>
                <Text style={styles.addDateBtnText}>+ Add Date</Text>
              </Pressable>
            </View>

            {screeningDates.map((d, idx) => (
              <View style={styles.referralDateRow} key={d.id}>
                <Text style={styles.referralDateLabel}>Screening {idx + 1}</Text>
                <Pressable
                  style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                  onPress={() => openDatePicker(d.id, d.value, 'screening')}
                >
                  <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                </Pressable>
                {screeningDates.length > 1 && (
                  <Pressable
                    style={styles.customFieldRemoveBtn}
                    onPress={() => removeScreeningDate(d.id)}
                  >
                    <Text style={styles.customFieldRemoveText}>X</Text>
                  </Pressable>
                )}
              </View>
            ))}

            {/* Only show referral dates section if NOT UR */}
            {!isUr && (
              <>
                <View style={styles.sectionHeaderRow}>
                  <Text style={styles.sectionHeaderTitle}>Referral Dates</Text>
                  <Pressable style={styles.addDateBtn} onPress={addReferralDate}>
                    <Text style={styles.addDateBtnText}>+ Add Date</Text>
                  </Pressable>
                </View>

                {referralDates.length === 0 && (
                  <Text style={{ color: '#9ca3af', fontSize: 14, fontStyle: 'italic', marginBottom: 8 }}>No referral dates added</Text>
                )}

                {referralDates.map((d, idx) => (
                  <View style={styles.referralDateRow} key={d.id}>
                    <Text style={styles.referralDateLabel}>Referral {idx + 1}</Text>
                    <Pressable
                      style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                      onPress={() => openDatePicker(d.id, d.value, 'referral')}
                    >
                      <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                    </Pressable>
                    <Pressable
                      style={styles.customFieldRemoveBtn}
                      onPress={() => removeReferralDate(d.id)}
                    >
                      <Text style={styles.customFieldRemoveText}>X</Text>
                    </Pressable>
                  </View>
                ))}
              </>
            )}

            {isUr && (
              <Text style={{ color: '#065f46', fontSize: 14, fontStyle: 'italic', marginBottom: 8 }}>
                (UR patient - no referral date needed)
              </Text>
            )}

            {datePicker.visible && (
              <DateTimePicker
                value={(() => {
                  const parts = datePicker.currentValue.split('/');
                  if (parts.length === 3) {
                    return new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1]));
                  }
                  return new Date();
                })()}
                mode="date"
                display={Platform.OS === 'ios' ? 'spinner' : 'default'}
                onChange={onDateChange}
              />
            )}

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Right Breast</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Suspicious"
                placeholderTextColor="#64748b"
                value={rightBreast}
                onChangeText={setRightBreast}
              />
            </View>

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Left Breast</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Suspicious"
                placeholderTextColor="#64748b"
                value={leftBreast}
                onChangeText={setLeftBreast}
              />
            </View>

            <View style={styles.quickStatusRow}>
              <Pressable
                style={[
                  styles.quickStatusBtn,
                  isUr && styles.quickStatusBtnActiveUr,
                ]}
                onPress={() => {
                  setRightBreast('UR');
                  setLeftBreast('UR');
                }}
              >
                <Text style={[
                  styles.quickStatusBtnText,
                  isUr && { color: '#065f46' }
                ]}>UR (Both)</Text>
              </Pressable>
              <Pressable
                style={[
                  styles.quickStatusBtn,
                  !isUr && (rightBreast !== '(empty)' || leftBreast !== '(empty)') && styles.quickStatusBtnActiveNotUr,
                ]}
                onPress={() => {
                  if (rightBreast === 'UR' || rightBreast === '(empty)') setRightBreast('Suspicious');
                  if (leftBreast === 'UR' || leftBreast === '(empty)') setLeftBreast('Suspicious');
                }}
              >
                <Text style={[
                  styles.quickStatusBtnText,
                  !isUr && (rightBreast !== '(empty)' || leftBreast !== '(empty)') && { color: '#92400e' }
                ]}>Suspicious</Text>
              </Pressable>
            </View>

            {/* All categorized fields */}
            {/* Demographics */}
            {editDemographics.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Patient Demographics</Text>
                {editDemographics.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* OB-GYNE History */}
            {(editMenstrualHistory.length > 0 || editPregnancyHistory.length > 0 || editContraceptiveUse.length > 0) && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>OB-GYNE History</Text>              
                {editMenstrualHistory.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Menstrual History</Text>
                    {editMenstrualHistory.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
                
                {editPregnancyHistory.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Pregnancy History</Text>
                    {editPregnancyHistory.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
                
                {editContraceptiveUse.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Contraceptive Use</Text>
                    {editContraceptiveUse.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
              </View>
            )}

            {/* Sexual History */}
            {editSexualHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Sexual History</Text>
                {editSexualHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Family & Social History */}
            {editFamilySocialHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Family & Social History</Text>
                {editFamilySocialHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Medical History */}
            {editMedicalHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Medical History</Text>
                {editMedicalHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Vital Signs */}
            {editVitalSigns.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Vital Signs</Text>
                {editVitalSigns.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Anthropometric */}
            {editAnthropometric.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Anthropometric</Text>
                {editAnthropometric.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Physical Examination */}
            {editPhysicalExam.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Physical Examination</Text>
                {editPhysicalExam.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Other */}
            {editOtherKeys.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Other Details</Text>
                {editOtherKeys.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            <Pressable style={[styles.saveBtn, { marginTop: 32 }]} onPress={handleSave}>
              <LinearGradient
                colors={['#db4278', '#b51f55']}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 0 }}
                style={StyleSheet.absoluteFillObject}
              />
              <Text style={styles.saveBtnText}>Save Changes</Text>
            </Pressable>
          </ScrollView>
        </KeyboardAvoidingView>
      </SafeAreaView>
    </Modal>
  );
}

function generateAdmitNumber(): string {
  return String(Math.floor(100000000 + Math.random() * 900000000));
}

function todayString(): string {
  const today = new Date();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  const yyyy = today.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
}

/** Split field keys into grouped display categories for edit/add forms */
function categorizeFieldKeys(keys: string[]) {
  const menstrualHistory: string[] = [];
  const pregnancyHistory: string[] = [];
  const contraceptiveUse: string[] = [];
  const sexualHistory: string[] = [];
  const familySocialHistory: string[] = [];
  const medicalHistory: string[] = [];
  const vitalSigns: string[] = [];
  const anthropometric: string[] = [];
  const physicalExam: string[] = [];
  const demographics: string[] = [];
  const other: string[] = [];

  // Group definitions – extended with variations
  const menstrualSet = new Set([
    'menarche', 'lmp', 'aog', 'menstrual bleeding pattern', 'no. of pads/day', 'no. of pads'
  ]);
  const pregnancySet = new Set([
    'pregnancy history', 'gtpal', 'age at first full term pregnancy'
  ]);
  const contraceptiveSet = new Set([
    'oral contraceptives use', 'oral contraceptive use', 'duration of usage of oral contraceptives'
  ]);
  const sexualSet = new Set([
    'age of first intercourse', 'no. of sexual partner', 'no. of sexual partners', 'spouse/partner/s'
  ]);
  const familySocialSet = new Set([
    'family history of cancer', 'smoking'
  ]);
  const medicalSet = new Set([
    'current medication', 'allergies', 'abdominal surgery',
    'history of previous cervical cancer screening',
    'history of abnormal vaginal discharge',
    'history of abnormal vaginal bleeding'
  ]);
  const vitalSet = new Set([
    'bp', 'temp', 'hr', 'rr', 'temperature', 'blood pressure', 'pulse rate', 'respiratory rate'
  ]);
  const anthropometricSet = new Set([
    'height (cm)', 'height', 'weight (kg)', 'weight', 'bmi'
  ]);
  const physicalExamSet = new Set([
    'skin', 'heent', 'chest and lungs', 'chest and luns', 'heart', 'abdomen'
  ]);
  const demographicSet = new Set([
    'id no. (hosp)', 'age', 'address', 'contact no.', 'civil status', 'no. of children'
  ]);

  keys.forEach((key) => {
    const lower = key.toLowerCase().trim();

    // --- Most specific first ---
    if (menstrualSet.has(lower) || [...menstrualSet].some((k) => lower.includes(k) && lower.length > 3)) {
      menstrualHistory.push(key);
    } else if (pregnancySet.has(lower) || [...pregnancySet].some((k) => lower.includes(k) && lower.length > 3)) {
      pregnancyHistory.push(key);
    } else if (contraceptiveSet.has(lower) || [...contraceptiveSet].some((k) => lower.includes(k) && lower.length > 3)) {
      contraceptiveUse.push(key);
    } else if (sexualSet.has(lower) || [...sexualSet].some((k) => lower.includes(k) && lower.length > 3)) {
      sexualHistory.push(key);
    } else if (familySocialSet.has(lower) || [...familySocialSet].some((k) => lower.includes(k) && lower.length > 3)) {
      familySocialHistory.push(key);
    } else if (medicalSet.has(lower) || [...medicalSet].some((k) => lower.includes(k) && lower.length > 3)) {
      medicalHistory.push(key);
    } else if (vitalSet.has(lower) || [...vitalSet].some((k) => lower.includes(k) && lower.length > 3)) {
      vitalSigns.push(key);
    } else if (anthropometricSet.has(lower) || [...anthropometricSet].some((k) => lower.includes(k) && lower.length > 3)) {
      anthropometric.push(key);
    } else if (physicalExamSet.has(lower) || [...physicalExamSet].some((k) => lower.includes(k) && lower.length > 3)) {
      physicalExam.push(key);
    } else if (demographicSet.has(lower) || [...demographicSet].some((k) => lower.includes(k) && lower.length > 3)) {
      demographics.push(key);
    } else {
      other.push(key);
    }
  });

  return { menstrualHistory, pregnancyHistory, contraceptiveUse, sexualHistory, familySocialHistory, medicalHistory, vitalSigns, anthropometric, physicalExam, demographics, other };
}

function checkDuplicateDate(
  name: string,
  newDateValues: string[],
  existingPatients: PatientRecord[] | undefined,
  onProceed: () => void,
  dateType: 'screening' | 'referral' = 'screening',
): boolean {
  if (!existingPatients || existingPatients.length === 0) {
    onProceed();
    return true;
  }

  const sameNamePatients = existingPatients.filter(
    (p) => p.patientName.toLowerCase().trim() === name.toLowerCase().trim()
  );

  for (const existing of sameNamePatients) {
    const existingDates = dateType === 'screening'
      ? existing.screeningDate.split(';').map((d) => d.trim()).filter(Boolean)
      : existing.date.split(';').map((d) => d.trim()).filter(Boolean);
    for (const newDate of newDateValues) {
      if (existingDates.includes(newDate)) {
        Alert.alert(
          'Duplicate Date Warning',
          `The ${dateType} date "${newDate}" already exists for patient "${existing.patientName}".\n\nDo you want to continue?`,
          [
            { text: 'Cancel', style: 'cancel' },
            { text: 'Continue', onPress: onProceed },
          ]
        );
        return false;
      }
    }
  }

  onProceed();
  return true;
}

function AddPatientModal({
  visible,
  onClose,
  onSave,
  existingPatients,
}: {
  visible: boolean;
  onClose: () => void;
  onSave: (patient: {
    patientName: string;
    date: string;
    screeningDate: string;
    rightBreastValue: string;
    leftBreastValue: string;
    breastValue: string;
    fields: Record<string, string>;
  }) => void;
  existingPatients?: PatientRecord[];
}) {
  const [name, setName] = useState('');
  const [screeningDates, setScreeningDates] = useState<Array<{ id: string; value: string }>>([{ id: 'init', value: todayString() }]);
  const [referralDates, setReferralDates] = useState<Array<{ id: string; value: string }>>([]);
  const [admitNumber, setAdmitNumber] = useState('');
  const [rightBreast, setRightBreast] = useState('UR');
  const [leftBreast, setLeftBreast] = useState('UR');
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const [customFields, setCustomFields] = useState<Array<{ id: string; label: string; value: string }>>([]);
  const [datePicker, setDatePicker] = useState<DatePickerState>({
    visible: false,
    mode: 'single',
    targetId: null,
    currentValue: '',
    dateType: 'screening',
  });

  const isUr = isPatientUr(rightBreast, leftBreast);

  // Use EXCEL_FIELD_KEYS merged with existing patient fields
  const fieldKeys = useMemo(() => {
    const baseKeys = EXCEL_FIELD_KEYS;
    if (existingPatients && existingPatients.length > 0) {
      const existingKeys = Object.keys(existingPatients[0].fields);
      const allKeys = Array.from(new Set([...baseKeys, ...existingKeys]));
      return allKeys.filter((key) => {
        const lowerKey = key.toLowerCase();
        return (
          !lowerKey.includes('name') &&
          !lowerKey.includes('date') &&
          !lowerKey.includes('breast') &&
          !lowerKey.includes('admit number') &&
          !lowerKey.includes('admit no') &&
          !lowerKey.includes('screening')
        );
      });
    }
    return baseKeys;
  }, [existingPatients]);

  useEffect(() => {
    if (visible) {
      setName('');
      setScreeningDates([{ id: 'init', value: todayString() }]);
      setReferralDates([]);
      setAdmitNumber(generateAdmitNumber());
      setRightBreast('UR');
      setLeftBreast('UR');
      const initialValues: Record<string, string> = {};
      fieldKeys.forEach((key) => { initialValues[key] = ''; });
      setFieldValues(initialValues);
      setCustomFields([]);
    }
  }, [visible, fieldKeys]);

  const addScreeningDate = () => {
    setScreeningDates((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), value: todayString() },
    ]);
  };

  const updateScreeningDate = (id: string, value: string) => {
    setScreeningDates((prev) => prev.map((d) => (d.id === id ? { ...d, value } : d)));
  };

  const removeScreeningDate = (id: string) => {
    setScreeningDates((prev) => prev.filter((d) => d.id !== id));
  };

  const addReferralDate = () => {
    setReferralDates((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), value: todayString() },
    ]);
  };

  const updateReferralDate = (id: string, value: string) => {
    setReferralDates((prev) => prev.map((d) => (d.id === id ? { ...d, value } : d)));
  };

  const removeReferralDate = (id: string) => {
    setReferralDates((prev) => prev.filter((d) => d.id !== id));
  };

  const openDatePicker = (id: string, currentValue: string, dateType: 'screening' | 'referral') => {
    setDatePicker({
      visible: true,
      mode: 'single',
      targetId: id,
      currentValue,
      dateType,
    });
  };

  const onDateChange = (event: any, selectedDate?: Date) => {
    if (selectedDate) {
      const mm = String(selectedDate.getMonth() + 1).padStart(2, '0');
      const dd = String(selectedDate.getDate()).padStart(2, '0');
      const yyyy = selectedDate.getFullYear();
      const formatted = `${mm}/${dd}/${yyyy}`;

      if (datePicker.targetId) {
        if (datePicker.dateType === 'screening') {
          updateScreeningDate(datePicker.targetId, formatted);
        } else {
          updateReferralDate(datePicker.targetId, formatted);
        }
      }
    }
    setDatePicker((prev) => ({ ...prev, visible: false }));
  };

  const handleAddField = () => {
    setCustomFields((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), label: '', value: '' },
    ]);
  };

  const handleUpdateCustomField = (id: string, key: 'label' | 'value', text: string) => {
    setCustomFields((prev) =>
      prev.map((item) => (item.id === id ? { ...item, [key]: text } : item)),
    );
  };

  const handleRemoveCustomField = (id: string) => {
    setCustomFields((prev) => prev.filter((item) => item.id !== id));
  };

  const proceedWithSave = () => {
    const combinedScreening = screeningDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryScreening = screeningDates[0]?.value.trim() || '(no date)';

    const combinedReferral = referralDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryReferral = referralDates[0]?.value.trim() || '(no date)';

    const finalFields: Record<string, string> = {};

    let nameKey = 'name';
    let dateKey = 'date';
    let screeningKey = 'Screening Date';
    let rightBreastKey = 'Right Breast';
    let leftBreastKey = 'Left Breast';
    let admitKey = 'Admit Number';

    if (existingPatients && existingPatients.length > 0) {
      const firstPatientKeys = Object.keys(existingPatients[0].fields);

      const foundNameKey = firstPatientKeys.find(
        (k) => k.toLowerCase() === 'name' || k.toLowerCase().includes('patient name'),
      );
      if (foundNameKey) nameKey = foundNameKey;

      const foundDateKey = firstPatientKeys.find((k) => k.toLowerCase() === 'date' || k.toLowerCase().includes('referral date'));
      if (foundDateKey) dateKey = foundDateKey;

      const foundScreeningKey = firstPatientKeys.find((k) => k.toLowerCase().includes('screening date'));
      if (foundScreeningKey) screeningKey = foundScreeningKey;

      const foundRightBreastKey = firstPatientKeys.find((k) => k.toLowerCase().includes('right breast'));
      if (foundRightBreastKey) rightBreastKey = foundRightBreastKey;

      const foundLeftBreastKey = firstPatientKeys.find((k) => k.toLowerCase().includes('left breast'));
      if (foundLeftBreastKey) leftBreastKey = foundLeftBreastKey;

      const foundAdmitKey = firstPatientKeys.find((k) => k.toLowerCase().includes('admit number') || k.toLowerCase().includes('admit no'));
      if (foundAdmitKey) admitKey = foundAdmitKey;

      firstPatientKeys.forEach((k) => {
        if (k === nameKey) {
          finalFields[k] = name.trim();
        } else if (k === dateKey) {
          finalFields[k] = combinedReferral;
        } else if (k === screeningKey) {
          finalFields[k] = combinedScreening;
        } else if (k === rightBreastKey) {
          finalFields[k] = rightBreast.trim();
        } else if (k === leftBreastKey) {
          finalFields[k] = leftBreast.trim();
        } else if (k === admitKey) {
          finalFields[k] = admitNumber.trim();
        } else {
          finalFields[k] = (fieldValues[k] || '').trim();
        }
      });
    } else {
      finalFields['Screening Date'] = combinedScreening;
      finalFields['Referral Date'] = combinedReferral;
      finalFields['Admit Number'] = admitNumber.trim();
      finalFields['name'] = name.trim();
      finalFields['Right Breast'] = rightBreast.trim();
      finalFields['Left Breast'] = leftBreast.trim();
      // Add all Excel fields
      fieldKeys.forEach((key) => {
        finalFields[key] = (fieldValues[key] || '').trim();
      });
    }

    // Ensure admit number is always present
    finalFields['Admit Number'] = admitNumber.trim();

    customFields.forEach((cf) => {
      if (cf.label.trim()) {
        finalFields[cf.label.trim()] = cf.value.trim();
      }
    });

    // Compute breastValue for display
    const rightIsUr = rightBreast.toUpperCase() === 'UR';
    const leftIsUr = leftBreast.toUpperCase() === 'UR';
    const isUrComputed = rightIsUr && leftIsUr;
    const breastValue = isUrComputed
      ? 'UR'
      : [
          !rightIsUr && rightBreast ? `R: ${rightBreast}` : '',
          !leftIsUr && leftBreast ? `L: ${leftBreast}` : '',
        ]
          .filter(Boolean)
          .join(' | ') || '(empty)';

    onSave({
      patientName: name.trim(),
      date: primaryReferral,
      screeningDate: primaryScreening,
      rightBreastValue: rightBreast.trim(),
      leftBreastValue: leftBreast.trim(),
      breastValue,
      fields: finalFields,
    });
  };

  const handleSave = () => {
    if (!name.trim()) {
      Alert.alert('Validation Error', 'Patient Name is required.');
      return;
    }

    const screeningDateValues = screeningDates.map((d) => d.value.trim()).filter(Boolean);
    if (screeningDateValues.length === 0) {
      Alert.alert('Validation Error', 'At least one Screening Date is required.');
      return;
    }

    const newScreeningValues = screeningDates.map((d) => d.value.trim()).filter(Boolean);
    checkDuplicateDate(name.trim(), newScreeningValues, existingPatients, proceedWithSave, 'screening');
  };

  const {
    menstrualHistory,
    pregnancyHistory,
    contraceptiveUse,
    sexualHistory,
    familySocialHistory,
    medicalHistory,
    vitalSigns,
    anthropometric,
    physicalExam,
    demographics,
    other,
  } = categorizeFieldKeys(fieldKeys);

  return (
    <Modal visible={visible} animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <KeyboardAvoidingView
          behavior={Platform.OS === 'ios' ? 'padding' : 'height'}
          style={{ flex: 1 }}
        >
          <View style={styles.modalHeaderRow}>
            <Text style={styles.modalHeaderTitle}>Add Patient</Text>
            <Pressable style={styles.modalCancelButton} onPress={onClose}>
              <Text style={styles.modalCancelButtonText}>Cancel</Text>
            </Pressable>
          </View>

          <ScrollView style={styles.modalScrollView} contentContainerStyle={styles.modalScrollContent}>
            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Patient Name *</Text>
              <TextInput
                style={styles.formInput}
                placeholder="Enter patient name"
                placeholderTextColor="#64748b"
                value={name}
                onChangeText={setName}
              />
            </View>

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Admit Number</Text>
              <View style={styles.admitNumberRow}>
                <TextInput
                  style={[styles.formInput, { flex: 1 }]}
                  placeholder="9-digit admit number"
                  placeholderTextColor="#64748b"
                  value={admitNumber}
                  onChangeText={setAdmitNumber}
                  keyboardType="numeric"
                  maxLength={9}
                />
                <Pressable
                  style={styles.regenBtn}
                  onPress={() => setAdmitNumber(generateAdmitNumber())}
                >
                  <Text style={styles.regenBtnText}>New</Text>
                </Pressable>
              </View>
            </View>

            <View style={styles.sectionHeaderRow}>
              <Text style={styles.sectionHeaderTitle}>Screening Dates *</Text>
              <Pressable style={styles.addDateBtn} onPress={addScreeningDate}>
                <Text style={styles.addDateBtnText}>+ Add Date</Text>
              </Pressable>
            </View>

            {screeningDates.map((d, idx) => (
              <View style={styles.referralDateRow} key={d.id}>
                <Text style={styles.referralDateLabel}>Screening {idx + 1}</Text>
                <Pressable
                  style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                  onPress={() => openDatePicker(d.id, d.value, 'screening')}
                >
                  <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                </Pressable>
                {screeningDates.length > 1 && (
                  <Pressable
                    style={styles.customFieldRemoveBtn}
                    onPress={() => removeScreeningDate(d.id)}
                  >
                    <Text style={styles.customFieldRemoveText}>X</Text>
                  </Pressable>
                )}
              </View>
            ))}

            {/* Only show referral dates section if NOT UR */}
            {!isUr && (
              <>
                <View style={styles.sectionHeaderRow}>
                  <Text style={styles.sectionHeaderTitle}>Referral Dates</Text>
                  <Pressable style={styles.addDateBtn} onPress={addReferralDate}>
                    <Text style={styles.addDateBtnText}>+ Add Date</Text>
                  </Pressable>
                </View>

                {referralDates.length === 0 && (
                  <Text style={{ color: '#9ca3af', fontSize: 14, fontStyle: 'italic', marginBottom: 8 }}>No referral dates added</Text>
                )}

                {referralDates.map((d, idx) => (
                  <View style={styles.referralDateRow} key={d.id}>
                    <Text style={styles.referralDateLabel}>Referral {idx + 1}</Text>
                    <Pressable
                      style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                      onPress={() => openDatePicker(d.id, d.value, 'referral')}
                    >
                      <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                    </Pressable>
                    <Pressable
                      style={styles.customFieldRemoveBtn}
                      onPress={() => removeReferralDate(d.id)}
                    >
                      <Text style={styles.customFieldRemoveText}>X</Text>
                    </Pressable>
                  </View>
                ))}
              </>
            )}

            {isUr && (
              <Text style={{ color: '#065f46', fontSize: 14, fontStyle: 'italic', marginBottom: 8 }}>
                (UR patient - no referral date needed)
              </Text>
            )}

            {datePicker.visible && (
              <DateTimePicker
                value={(() => {
                  const parts = datePicker.currentValue.split('/');
                  if (parts.length === 3) {
                    return new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1]));
                  }
                  return new Date();
                })()}
                mode="date"
                display={Platform.OS === 'ios' ? 'spinner' : 'default'}
                onChange={onDateChange}
              />
            )}

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Right Breast *</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Suspicious"
                placeholderTextColor="#64748b"
                value={rightBreast}
                onChangeText={setRightBreast}
              />
            </View>

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Left Breast *</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Suspicious"
                placeholderTextColor="#64748b"
                value={leftBreast}
                onChangeText={setLeftBreast}
              />
            </View>

            <View style={styles.quickStatusRow}>
              <Pressable
                style={[
                  styles.quickStatusBtn,
                  isUr && styles.quickStatusBtnActiveUr,
                ]}
                onPress={() => {
                  setRightBreast('UR');
                  setLeftBreast('UR');
                }}
              >
                <Text style={[
                  styles.quickStatusBtnText,
                  isUr && { color: '#065f46' }
                ]}>UR (Both)</Text>
              </Pressable>
              <Pressable
                style={[
                  styles.quickStatusBtn,
                  !isUr && (rightBreast !== 'UR' || leftBreast !== 'UR') && styles.quickStatusBtnActiveNotUr,
                ]}
                onPress={() => {
                  if (rightBreast === 'UR' || rightBreast === '') setRightBreast('Suspicious');
                  if (leftBreast === 'UR' || leftBreast === '') setLeftBreast('Suspicious');
                }}
              >
                <Text style={[
                  styles.quickStatusBtnText,
                  !isUr && (rightBreast !== 'UR' || leftBreast !== 'UR') && { color: '#92400e' }
                ]}>Suspicious</Text>
              </Pressable>
            </View>

            {/* All categorized fields */}
            {/* Demographics */}
            {demographics.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Patient Demographics</Text>
                {demographics.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* OB-GYNE History */}
            {(menstrualHistory.length > 0 || pregnancyHistory.length > 0 || contraceptiveUse.length > 0) && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>OB-GYNE History</Text>
                
                {menstrualHistory.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Menstrual History</Text>
                    {menstrualHistory.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
                
                {pregnancyHistory.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Pregnancy History</Text>
                    {pregnancyHistory.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
                
                {contraceptiveUse.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Contraceptive Use</Text>
                    {contraceptiveUse.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
              </View>
            )}

            {/* Sexual History */}
            {sexualHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Sexual History</Text>
                {sexualHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Family & Social History */}
            {familySocialHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Family & Social History</Text>
                {familySocialHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Medical History */}
            {medicalHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Medical History</Text>
                {medicalHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Vital Signs */}
            {vitalSigns.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Vital Signs</Text>
                {vitalSigns.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Anthropometric */}
            {anthropometric.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Anthropometric</Text>
                {anthropometric.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Physical Examination */}
            {physicalExam.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Physical Examination</Text>
                {physicalExam.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Other */}
            {other.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Other Details</Text>
                {other.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {customFields.length > 0 && (
              <View style={styles.sectionHeaderRow}>
                <Text style={styles.sectionHeaderTitle}>Custom Fields</Text>
              </View>
            )}

            {customFields.map((cf) => (
              <View style={styles.customFieldRow} key={cf.id}>
                <TextInput
                  style={[styles.formInput, styles.customFieldLabelInput]}
                  placeholder="Label"
                  placeholderTextColor="#64748b"
                  value={cf.label}
                  onChangeText={(text) => handleUpdateCustomField(cf.id, 'label', text)}
                />
                <TextInput
                  style={[styles.formInput, styles.customFieldValueInput]}
                  placeholder="Value"
                  placeholderTextColor="#64748b"
                  value={cf.value}
                  onChangeText={(text) => handleUpdateCustomField(cf.id, 'value', text)}
                />
                <Pressable
                  style={styles.customFieldRemoveBtn}
                  onPress={() => handleRemoveCustomField(cf.id)}
                >
                  <Text style={styles.customFieldRemoveText}>X</Text>
                </Pressable>
              </View>
            ))}

            <Pressable style={styles.addFieldBtn} onPress={handleAddField}>
              <Text style={styles.addFieldBtnText}>+ Add Custom Field</Text>
            </Pressable>

            <Pressable style={styles.saveBtn} onPress={handleSave}>
              <LinearGradient
                colors={['#db4278', '#b51f55']}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 0 }}
                style={StyleSheet.absoluteFillObject}
              />
              <Text style={styles.saveBtnText}>Save Patient Record</Text>
            </Pressable>
          </ScrollView>
        </KeyboardAvoidingView>
      </SafeAreaView>
    </Modal>
  );
}

function AddRecordForPatientModal({
  patientName,
  onClose,
  onSave,
  existingPatients,
}: {
  patientName: string;
  onClose: () => void;
  onSave: (patient: {
    patientName: string;
    date: string;
    screeningDate: string;
    rightBreastValue: string;
    leftBreastValue: string;
    breastValue: string;
    fields: Record<string, string>;
  }) => void;
  existingPatients?: PatientRecord[];
}) {
  const [name, setName] = useState(patientName);
  const [screeningDates, setScreeningDates] = useState<Array<{ id: string; value: string }>>([{ id: 'init', value: todayString() }]);
  const [referralDates, setReferralDates] = useState<Array<{ id: string; value: string }>>([]);
  const [admitNumber, setAdmitNumber] = useState(generateAdmitNumber());
  const [rightBreast, setRightBreast] = useState('UR');
  const [leftBreast, setLeftBreast] = useState('UR');
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const [customFields, setCustomFields] = useState<Array<{ id: string; label: string; value: string }>>([]);
  const [datePicker, setDatePicker] = useState<DatePickerState>({
    visible: false,
    mode: 'single',
    targetId: null,
    currentValue: '',
    dateType: 'screening',
  });

  const isUr = isPatientUr(rightBreast, leftBreast);

  // Use EXCEL_FIELD_KEYS merged with existing patient fields
  const fieldKeys = useMemo(() => {
    const baseKeys = EXCEL_FIELD_KEYS;
    if (existingPatients && existingPatients.length > 0) {
      const existingKeys = Object.keys(existingPatients[0].fields);
      const allKeys = Array.from(new Set([...baseKeys, ...existingKeys]));
      return allKeys.filter((key) => {
        const lowerKey = key.toLowerCase();
        return (
          !lowerKey.includes('name') &&
          !lowerKey.includes('date') &&
          !lowerKey.includes('breast') &&
          !lowerKey.includes('admit number') &&
          !lowerKey.includes('admit no') &&
          !lowerKey.includes('screening')
        );
      });
    }
    return baseKeys;
  }, [existingPatients]);

  // Pre-fills Breast, custom fields, and all other details from the patient's latest record
  useEffect(() => {
    setName(patientName);

    const patientRecords = existingPatients
      ?.filter(
        (p) => p.patientName.toLowerCase().trim() === patientName.toLowerCase().trim()
      )
      .sort((a, b) => b.rowIndex - a.rowIndex);

    const latest = patientRecords?.[0];

    if (latest) {
      setScreeningDates([{ id: 'init', value: todayString() }]);
      setReferralDates([]);
      setAdmitNumber(generateAdmitNumber());
      setRightBreast(latest.rightBreastValue || 'UR');
      setLeftBreast(latest.leftBreastValue || 'UR');

      const initialValues: Record<string, string> = {};
      fieldKeys.forEach((key) => {
        initialValues[key] = latest.fields[key] || '';
      });
      setFieldValues(initialValues);
      setCustomFields([]);
    } else {
      setScreeningDates([{ id: 'init', value: todayString() }]);
      setReferralDates([]);
      setAdmitNumber(generateAdmitNumber());
      setRightBreast('UR');
      setLeftBreast('UR');
      const initialValues: Record<string, string> = {};
      fieldKeys.forEach((key) => { initialValues[key] = ''; });
      setFieldValues(initialValues);
      setCustomFields([]);
    }
  }, [patientName, fieldKeys, existingPatients]);

  const addScreeningDate = () => {
    setScreeningDates((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), value: todayString() },
    ]);
  };

  const updateScreeningDate = (id: string, value: string) => {
    setScreeningDates((prev) => prev.map((d) => (d.id === id ? { ...d, value } : d)));
  };

  const removeScreeningDate = (id: string) => {
    setScreeningDates((prev) => prev.filter((d) => d.id !== id));
  };

  const addReferralDate = () => {
    setReferralDates((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), value: todayString() },
    ]);
  };

  const updateReferralDate = (id: string, value: string) => {
    setReferralDates((prev) => prev.map((d) => (d.id === id ? { ...d, value } : d)));
  };

  const removeReferralDate = (id: string) => {
    setReferralDates((prev) => prev.filter((d) => d.id !== id));
  };

  const openDatePicker = (id: string, currentValue: string, dateType: 'screening' | 'referral') => {
    setDatePicker({
      visible: true,
      mode: 'single',
      targetId: id,
      currentValue,
      dateType,
    });
  };

  const onDateChange = (event: any, selectedDate?: Date) => {
    if (selectedDate) {
      const mm = String(selectedDate.getMonth() + 1).padStart(2, '0');
      const dd = String(selectedDate.getDate()).padStart(2, '0');
      const yyyy = selectedDate.getFullYear();
      const formatted = `${mm}/${dd}/${yyyy}`;

      if (datePicker.targetId) {
        if (datePicker.dateType === 'screening') {
          updateScreeningDate(datePicker.targetId, formatted);
        } else {
          updateReferralDate(datePicker.targetId, formatted);
        }
      }
    }
    setDatePicker((prev) => ({ ...prev, visible: false }));
  };

  const handleAddField = () => {
    setCustomFields((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), label: '', value: '' },
    ]);
  };

  const handleUpdateCustomField = (id: string, key: 'label' | 'value', text: string) => {
    setCustomFields((prev) =>
      prev.map((item) => (item.id === id ? { ...item, [key]: text } : item)),
    );
  };

  const handleRemoveCustomField = (id: string) => {
    setCustomFields((prev) => prev.filter((item) => item.id !== id));
  };

  const proceedWithSave = () => {
    const combinedScreening = screeningDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryScreening = screeningDates[0]?.value.trim() || '(no date)';

    const combinedReferral = referralDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryReferral = referralDates[0]?.value.trim() || '(no date)';

    const finalFields: Record<string, string> = {};

    let nameKey = 'name';
    let dateKey = 'date';
    let screeningKey = 'Screening Date';
    let rightBreastKey = 'Right Breast';
    let leftBreastKey = 'Left Breast';
    let admitKey = 'Admit Number';

    if (existingPatients && existingPatients.length > 0) {
      const firstPatientKeys = Object.keys(existingPatients[0].fields);

      const foundNameKey = firstPatientKeys.find(
        (k) => k.toLowerCase() === 'name' || k.toLowerCase().includes('patient name'),
      );
      if (foundNameKey) nameKey = foundNameKey;

      const foundDateKey = firstPatientKeys.find((k) => k.toLowerCase() === 'date' || k.toLowerCase().includes('referral date'));
      if (foundDateKey) dateKey = foundDateKey;

      const foundScreeningKey = firstPatientKeys.find((k) => k.toLowerCase().includes('screening date'));
      if (foundScreeningKey) screeningKey = foundScreeningKey;

      const foundRightBreastKey = firstPatientKeys.find((k) => k.toLowerCase().includes('right breast'));
      if (foundRightBreastKey) rightBreastKey = foundRightBreastKey;

      const foundLeftBreastKey = firstPatientKeys.find((k) => k.toLowerCase().includes('left breast'));
      if (foundLeftBreastKey) leftBreastKey = foundLeftBreastKey;

      const foundAdmitKey = firstPatientKeys.find((k) => k.toLowerCase().includes('admit number') || k.toLowerCase().includes('admit no'));
      if (foundAdmitKey) admitKey = foundAdmitKey;

      firstPatientKeys.forEach((k) => {
        if (k === nameKey) {
          finalFields[k] = name.trim();
        } else if (k === dateKey) {
          finalFields[k] = combinedReferral;
        } else if (k === screeningKey) {
          finalFields[k] = combinedScreening;
        } else if (k === rightBreastKey) {
          finalFields[k] = rightBreast.trim();
        } else if (k === leftBreastKey) {
          finalFields[k] = leftBreast.trim();
        } else if (k === admitKey) {
          finalFields[k] = admitNumber.trim();
        } else {
          finalFields[k] = (fieldValues[k] || '').trim();
        }
      });
    } else {
      finalFields['Screening Date'] = combinedScreening;
      finalFields['Referral Date'] = combinedReferral;
      finalFields['Admit Number'] = admitNumber.trim();
      finalFields['name'] = name.trim();
      finalFields['Right Breast'] = rightBreast.trim();
      finalFields['Left Breast'] = leftBreast.trim();
      fieldKeys.forEach((key) => {
        finalFields[key] = (fieldValues[key] || '').trim();
      });
    }

    finalFields['Admit Number'] = admitNumber.trim();

    customFields.forEach((cf) => {
      if (cf.label.trim()) {
        finalFields[cf.label.trim()] = cf.value.trim();
      }
    });

    // Compute breastValue for display
    const rightIsUr = rightBreast.toUpperCase() === 'UR';
    const leftIsUr = leftBreast.toUpperCase() === 'UR';
    const isUrComputed = rightIsUr && leftIsUr;
    const breastValue = isUrComputed
      ? 'UR'
      : [
          !rightIsUr && rightBreast ? `R: ${rightBreast}` : '',
          !leftIsUr && leftBreast ? `L: ${leftBreast}` : '',
        ]
          .filter(Boolean)
          .join(' | ') || '(empty)';

    onSave({
      patientName: name.trim(),
      date: primaryReferral,
      screeningDate: primaryScreening,
      rightBreastValue: rightBreast.trim(),
      leftBreastValue: leftBreast.trim(),
      breastValue,
      fields: finalFields,
    });
  };

  const handleSave = () => {
    if (!name.trim()) {
      Alert.alert('Validation Error', 'Patient Name is required.');
      return;
    }

    const screeningDateValues = screeningDates.map((d) => d.value.trim()).filter(Boolean);
    if (screeningDateValues.length === 0) {
      Alert.alert('Validation Error', 'At least one Screening Date is required.');
      return;
    }

    const newScreeningValues = screeningDates.map((d) => d.value.trim()).filter(Boolean);
    checkDuplicateDate(name.trim(), newScreeningValues, existingPatients, proceedWithSave, 'screening');
  };

  const {
    menstrualHistory,
    pregnancyHistory,
    contraceptiveUse,
    sexualHistory,
    familySocialHistory,
    medicalHistory,
    vitalSigns,
    anthropometric,
    physicalExam,
    demographics,
    other,
  } = categorizeFieldKeys(fieldKeys);

  return (
    <Modal visible animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <KeyboardAvoidingView
          behavior={Platform.OS === 'ios' ? 'padding' : 'height'}
          style={{ flex: 1 }}
        >
          <View style={styles.modalHeaderRow}>
            <Text style={styles.modalHeaderTitle}>Add Record for {patientName}</Text>
            <Pressable style={styles.modalCancelButton} onPress={onClose}>
              <Text style={styles.modalCancelButtonText}>Cancel</Text>
            </Pressable>
          </View>

          <ScrollView style={styles.modalScrollView} contentContainerStyle={styles.modalScrollContent}>
            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Patient Name *</Text>
              <TextInput
                style={styles.formInput}
                placeholder="Enter patient name"
                placeholderTextColor="#64748b"
                value={name}
                onChangeText={setName}
              />
            </View>

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Admit Number</Text>
              <View style={styles.admitNumberRow}>
                <TextInput
                  style={[styles.formInput, { flex: 1 }]}
                  placeholder="9-digit admit number"
                  placeholderTextColor="#64748b"
                  value={admitNumber}
                  onChangeText={setAdmitNumber}
                  keyboardType="numeric"
                  maxLength={9}
                />
                <Pressable
                  style={styles.regenBtn}
                  onPress={() => setAdmitNumber(generateAdmitNumber())}
                >
                  <Text style={styles.regenBtnText}>New</Text>
                </Pressable>
              </View>
            </View>

            <View style={styles.sectionHeaderRow}>
              <Text style={styles.sectionHeaderTitle}>Screening Dates *</Text>
              <Pressable style={styles.addDateBtn} onPress={addScreeningDate}>
                <Text style={styles.addDateBtnText}>+ Add Date</Text>
              </Pressable>
            </View>

            {screeningDates.map((d, idx) => (
              <View style={styles.referralDateRow} key={d.id}>
                <Text style={styles.referralDateLabel}>Screening {idx + 1}</Text>
                <Pressable
                  style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                  onPress={() => openDatePicker(d.id, d.value, 'screening')}
                >
                  <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                </Pressable>
                {screeningDates.length > 1 && (
                  <Pressable
                    style={styles.customFieldRemoveBtn}
                    onPress={() => removeScreeningDate(d.id)}
                  >
                    <Text style={styles.customFieldRemoveText}>X</Text>
                  </Pressable>
                )}
              </View>
            ))}

            {/* Only show referral dates section if NOT UR */}
            {!isUr && (
              <>
                <View style={styles.sectionHeaderRow}>
                  <Text style={styles.sectionHeaderTitle}>Referral Dates</Text>
                  <Pressable style={styles.addDateBtn} onPress={addReferralDate}>
                    <Text style={styles.addDateBtnText}>+ Add Date</Text>
                  </Pressable>
                </View>

                {referralDates.length === 0 && (
                  <Text style={{ color: '#9ca3af', fontSize: 14, fontStyle: 'italic', marginBottom: 8 }}>No referral dates added</Text>
                )}

                {referralDates.map((d, idx) => (
                  <View style={styles.referralDateRow} key={d.id}>
                    <Text style={styles.referralDateLabel}>Referral {idx + 1}</Text>
                    <Pressable
                      style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                      onPress={() => openDatePicker(d.id, d.value, 'referral')}
                    >
                      <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                    </Pressable>
                    <Pressable
                      style={styles.customFieldRemoveBtn}
                      onPress={() => removeReferralDate(d.id)}
                    >
                      <Text style={styles.customFieldRemoveText}>X</Text>
                    </Pressable>
                  </View>
                ))}
              </>
            )}

            {isUr && (
              <Text style={{ color: '#065f46', fontSize: 14, fontStyle: 'italic', marginBottom: 8 }}>
                (UR patient - no referral date needed)
              </Text>
            )}

            {datePicker.visible && (
              <DateTimePicker
                value={(() => {
                  const parts = datePicker.currentValue.split('/');
                  if (parts.length === 3) {
                    return new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1]));
                  }
                  return new Date();
                })()}
                mode="date"
                display={Platform.OS === 'ios' ? 'spinner' : 'default'}
                onChange={onDateChange}
              />
            )}

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Right Breast *</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Suspicious"
                placeholderTextColor="#64748b"
                value={rightBreast}
                onChangeText={setRightBreast}
              />
            </View>

            <View style={styles.formGroup}>
              <Text style={styles.formLabel}>Left Breast *</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Suspicious"
                placeholderTextColor="#64748b"
                value={leftBreast}
                onChangeText={setLeftBreast}
              />
            </View>

            <View style={styles.quickStatusRow}>
              <Pressable
                style={[
                  styles.quickStatusBtn,
                  isUr && styles.quickStatusBtnActiveUr,
                ]}
                onPress={() => {
                  setRightBreast('UR');
                  setLeftBreast('UR');
                }}
              >
                <Text style={[
                  styles.quickStatusBtnText,
                  isUr && { color: '#065f46' }
                ]}>UR (Both)</Text>
              </Pressable>
              <Pressable
                style={[
                  styles.quickStatusBtn,
                  !isUr && (rightBreast !== 'UR' || leftBreast !== 'UR') && styles.quickStatusBtnActiveNotUr,
                ]}
                onPress={() => {
                  if (rightBreast === 'UR' || rightBreast === '') setRightBreast('Suspicious');
                  if (leftBreast === 'UR' || leftBreast === '') setLeftBreast('Suspicious');
                }}
              >
                <Text style={[
                  styles.quickStatusBtnText,
                  !isUr && (rightBreast !== 'UR' || leftBreast !== 'UR') && { color: '#92400e' }
                ]}>Suspicious</Text>
              </Pressable>
            </View>

            {/* All categorized fields */}
            {/* Demographics */}
            {demographics.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Patient Demographics</Text>
                {demographics.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* OB-GYNE History */}
            {(menstrualHistory.length > 0 || pregnancyHistory.length > 0 || contraceptiveUse.length > 0) && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>OB-GYNE History</Text>
                
                {menstrualHistory.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Menstrual History</Text>
                    {menstrualHistory.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
                
                {pregnancyHistory.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Pregnancy History</Text>
                    {pregnancyHistory.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
                
                {contraceptiveUse.length > 0 && (
                  <>
                    <Text style={styles.categorySubLabel}>Contraceptive Use</Text>
                    {contraceptiveUse.map((key) => (
                      <View style={styles.formGroup} key={key}>
                        <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                        <TextInput
                          style={styles.formInput}
                          placeholder={`Enter ${key.toLowerCase()}`}
                          placeholderTextColor="#64748b"
                          value={fieldValues[key] || ''}
                          onChangeText={(text) =>
                            setFieldValues((prev) => ({ ...prev, [key]: text }))
                          }
                        />
                      </View>
                    ))}
                  </>
                )}
              </View>
            )}

            {/* Sexual History */}
            {sexualHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Sexual History</Text>
                {sexualHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Family & Social History */}
            {familySocialHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Family & Social History</Text>
                {familySocialHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Medical History */}
            {medicalHistory.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Medical History</Text>
                {medicalHistory.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Vital Signs */}
            {vitalSigns.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Vital Signs</Text>
                {vitalSigns.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Anthropometric */}
            {anthropometric.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Anthropometric</Text>
                {anthropometric.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Physical Examination */}
            {physicalExam.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Physical Examination</Text>
                {physicalExam.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {/* Other */}
            {other.length > 0 && (
              <View style={styles.groupBox}>
                <Text style={styles.groupBoxTitle}>Other Details</Text>
                {other.map((key) => (
                  <View style={styles.formGroup} key={key}>
                    <Text style={styles.formLabel}>{getDisplayLabel(key)}</Text>
                    <TextInput
                      style={styles.formInput}
                      placeholder={`Enter ${key.toLowerCase()}`}
                      placeholderTextColor="#64748b"
                      value={fieldValues[key] || ''}
                      onChangeText={(text) =>
                        setFieldValues((prev) => ({ ...prev, [key]: text }))
                      }
                    />
                  </View>
                ))}
              </View>
            )}

            {customFields.length > 0 && (
              <View style={styles.sectionHeaderRow}>
                <Text style={styles.sectionHeaderTitle}>Custom Fields</Text>
              </View>
            )}

            {customFields.map((cf) => (
              <View style={styles.customFieldRow} key={cf.id}>
                <TextInput
                  style={[styles.formInput, styles.customFieldLabelInput]}
                  placeholder="Label"
                  placeholderTextColor="#64748b"
                  value={cf.label}
                  onChangeText={(text) => handleUpdateCustomField(cf.id, 'label', text)}
                />
                <TextInput
                  style={[styles.formInput, styles.customFieldValueInput]}
                  placeholder="Value"
                  placeholderTextColor="#64748b"
                  value={cf.value}
                  onChangeText={(text) => handleUpdateCustomField(cf.id, 'value', text)}
                />
                <Pressable
                  style={styles.customFieldRemoveBtn}
                  onPress={() => handleRemoveCustomField(cf.id)}
                >
                  <Text style={styles.customFieldRemoveText}>X</Text>
                </Pressable>
              </View>
            ))}

            <Pressable style={styles.addFieldBtn} onPress={handleAddField}>
              <Text style={styles.addFieldBtnText}>+ Add Custom Field</Text>
            </Pressable>

            <Pressable style={styles.saveBtn} onPress={handleSave}>
              <LinearGradient
                colors={['#db4278', '#b51f55']}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 0 }}
                style={StyleSheet.absoluteFillObject}
              />
              <Text style={styles.saveBtnText}>Save New Record</Text>
            </Pressable>
          </ScrollView>
        </KeyboardAvoidingView>
      </SafeAreaView>
    </Modal>
  );
}

function SidebarModal({
  visible,
  onClose,
  onAbout,
  onStaffs,
  onStatusBoard,
  onImport,
  onClearData,
  hasData,
}: {
  visible: boolean;
  onClose: () => void;
  onAbout: () => void;
  onStaffs: () => void;
  onStatusBoard: () => void;
  onImport: () => void;
  onClearData: () => void;
  hasData: boolean;
}) {
  return (
    <Modal visible={visible} animationType="fade" transparent onRequestClose={onClose}>
      <View style={styles.sidebarOverlay}>
        <Pressable style={styles.sidebarBackdrop} onPress={onClose} />
        <View style={styles.sidebarCard}>
          <View style={styles.sidebarHeader}>
            <Image source={teamSilayanLogo} style={styles.sidebarLogo} resizeMode="contain" />
            <Text style={styles.sidebarTitle}>Menu</Text>
          </View>

          <Pressable style={styles.sidebarItem} onPress={onAbout}>
            <Image
              source={require('./assets/information.png')}
              style={styles.sidebarItemIconImage}
              resizeMode="contain"
            />
            <Text style={styles.sidebarItemText}>About</Text>
          </Pressable>

          <Pressable style={styles.sidebarItem} onPress={onStaffs}>
          <Image
              source={require('./assets/team.png')}
              style={styles.sidebarItemIconImage}
              resizeMode="contain"
            />
            <Text style={styles.sidebarItemText}>Staffs</Text>
          </Pressable>

          <Pressable style={styles.sidebarItem} onPress={onStatusBoard}>
            <View style={styles.sidebarItemIconContainer}>
              <Text style={styles.sidebarItemIconText}>📋</Text>
            </View>
            <Text style={styles.sidebarItemText}>Status Board</Text>
          </Pressable>

          <Pressable style={styles.sidebarItem} onPress={() => {
            onClose();
            Alert.alert(
              'Import Excel',
              'This will import patient records from an Excel file. New records will be merged with your existing data.\n\nDo you want to continue?',
              [
                { text: 'Cancel', style: 'cancel' },
                { text: 'Continue', onPress: onImport },
              ]
            );
          }}>
            <View style={styles.sidebarItemIconContainer}>
              <Text style={styles.sidebarItemIconText}>📂</Text>
            </View>
            <Text style={styles.sidebarItemText}>Import Excel</Text>
          </Pressable>

          {hasData && (
            <Pressable
              style={[styles.sidebarItem, styles.sidebarItemDanger]}
              onPress={() => {
                onClose();
                onClearData();
              }}
            >
              <View style={styles.sidebarItemIconContainer}>
                <Text style={styles.sidebarItemIconDanger}>🗑</Text>
              </View>
              <Text style={styles.sidebarItemTextDanger}>Clear All Data</Text>
            </Pressable>
          )}

          <Pressable style={styles.sidebarCloseBtn} onPress={onClose}>
            <LinearGradient
              colors={['#db4278', '#b51f55']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 0 }}
              style={StyleSheet.absoluteFillObject}
            />
            <Text style={styles.sidebarCloseBtnText}>Close Menu</Text>
          </Pressable>
        </View>
      </View>
    </Modal>
  );
}

function StaffsModal({
  visible,
  onClose,
}: {
  visible: boolean;
  onClose: () => void;
}) {
  const [fullscreenImage, setFullscreenImage] = useState<any>(null);
  const [fullscreenName, setFullscreenName] = useState<string>('');

  const staffList = [
    { name: 'Jonathan Atendido', image: require('./assets/jonathan atendido.png'), role: 'BS in Biology Major in Medical Biology', location: 'Tetuan, Zamboanga City' },
    { name: 'Vilmardiya Kawthar Aman', image: require('./assets/vilmardiya aman.png'), role: 'BS in Biology Major in Medical Biology', location: 'Lamitan City, Basilan' },
    { name: 'Romar Cristian Ang', image: require('./assets/romar ang.png'), role: 'BS in Biology Major in Medical Biology', location: 'Tetuan, Zamboanga City' },
    { name: 'Alleah Cobong', image: require('./assets/alleah cobong.png'), role: 'BS Biology Major in Animal Biology', location: 'Pugaan, Iligan City' },
    { name: 'Nur-Jannah Godoy', image: require('./assets/nur-jannah godoy.png'), role: 'BS in Medical Laboratory Science', location: 'Mati City, Davao Oriental' },
    { name: 'Aneeka Laja', image: require('./assets/aneeka laja.png'), role: 'BS in Medical Technology', location: 'Siasi, Sulu' },
    { name: 'Sittie Hafizah Maguindra', image: require('./assets/sitti hafizah maguindra.png'), role: 'BS in Psychology', location: 'Pagadian City, Zamboanga del Sur' },
    { name: 'Shaira Mukaram', image: require('./assets/shaira mukaram.png'), role: 'BS in Medical Technology', location: 'Rio Hondo, Zamboanga City' },
    { name: 'Nurun Nahla Reyes', image: require('./assets/nurun nahla reyes.png'), role: 'BS in Medical Technology', location: 'Jolo, Sulu' },
    { name: 'Rasdin Sariul', image: require('./assets/rasdin sariul.png'), role: 'BS in Nursing', location: 'Cotabato City' },
    { name: 'Marwin Tubat', image: require('./assets/marwin tubat.png'), role: 'BS in Medical Technology', location: 'Sultan Naga Dimaporo, Lanao del Norte' },
  ];

  const handleImagePress = (staff: typeof staffList[0]) => {
    setFullscreenImage(staff.image);
    setFullscreenName(staff.name);
  };

  return (
    <Modal visible={visible} animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <View style={styles.modalHeaderRow}>
          <Text style={styles.modalHeaderTitle}>Staffs</Text>
          <Pressable style={styles.modalCancelButton} onPress={onClose}>
            <Text style={styles.modalCancelButtonText}>Close</Text>
          </Pressable>
        </View>

        <ScrollView style={styles.modalScrollView} contentContainerStyle={styles.modalScrollContent}>
          <Text style={styles.staffsSubtitle}>Team Silayan</Text>
          {staffList.map((staff, idx) => (
            <View key={idx} style={styles.staffCard}>
              <Pressable onPress={() => handleImagePress(staff)}>
                <Image source={staff.image} style={styles.staffAvatarImage} resizeMode="cover" />
              </Pressable>
              <View style={styles.staffInfo}>
                <Text style={styles.staffName}>{staff.name}</Text>
                <Text style={styles.staffRole}>{staff.role}</Text>
                <Text style={styles.staffDept}>{staff.location}</Text>
              </View>
            </View>
          ))}
        </ScrollView>
      </SafeAreaView>

      {/* Fullscreen Image Modal */}
      <Modal
        visible={fullscreenImage !== null}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setFullscreenImage(null)}
      >
        <View style={styles.staffFullscreenOverlay}>
          <Pressable
            style={styles.staffFullscreenCloseBtn}
            onPress={() => setFullscreenImage(null)}
          >
            <Text style={styles.staffFullscreenCloseText}>✕</Text>
          </Pressable>
          <Image
            source={fullscreenImage}
            style={styles.staffFullscreenImage}
            resizeMode="contain"
          />
          {fullscreenName ? (
            <Text style={styles.staffFullscreenName}>{fullscreenName}</Text>
          ) : null}
        </View>
      </Modal>
    </Modal>
  );
}

function StatusBoardModal({
  visible,
  onClose,
  patientGroups,
  onPatientPress,
}: {
  visible: boolean;
  onClose: () => void;
  patientGroups: PatientGroup[];
  onPatientPress: (patientName: string) => void;
}) {
  const [sortMode, setSortMode] = useState<'name' | 'status'>('name');
  const [sortAsc, setSortAsc] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');

  const sortedGroups = useMemo(() => {
    const sorted = [...patientGroups];
    sorted.sort((a, b) => {
      let cmp = 0;
      if (sortMode === 'name') {
        cmp = a.name.localeCompare(b.name);
      } else if (sortMode === 'status') {
        const statusA = a.records[a.records.length - 1]?.isUr ? 1 : 0;
        const statusB = b.records[b.records.length - 1]?.isUr ? 1 : 0;
        cmp = statusA - statusB;
      }
      return sortAsc ? cmp : -cmp;
    });
    return sorted;
  }, [patientGroups, sortMode, sortAsc]);

  const filteredGroups = useMemo(() => {
    if (!searchQuery.trim()) return sortedGroups;
    const q = searchQuery.toLowerCase().trim();
    return sortedGroups.filter((g) => g.name.toLowerCase().includes(q));
  }, [sortedGroups, searchQuery]);

  return (
    <Modal visible={visible} animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <View style={styles.modalHeaderRow}>
          <Text style={styles.modalHeaderTitle}>Status Board</Text>
          <Pressable style={styles.modalCancelButton} onPress={onClose}>
            <Text style={styles.modalCancelButtonText}>Close</Text>
          </Pressable>
        </View>

        <ScrollView style={styles.modalScrollView} contentContainerStyle={styles.modalScrollContent}>
          <View style={styles.searchRow}>
            <TextInput
              style={styles.searchInput}
              placeholder="Search patient name..."
              placeholderTextColor="#9ca3af"
              value={searchQuery}
              onChangeText={setSearchQuery}
            />
            {searchQuery.length > 0 && (
              <Pressable style={styles.searchClearBtn} onPress={() => setSearchQuery('')}>
                <Text style={styles.searchClearBtnText}>✕</Text>
              </Pressable>
            )}
          </View>

          <View style={styles.statusBoardListCard}>
            <View style={styles.statusBoardListHeader}>
              <Pressable style={[styles.statusBoardHeaderBtn, { flex: 1.5 }]} onPress={() => {
                if (sortMode === 'name') setSortAsc(!sortAsc);
                else { setSortMode('name'); setSortAsc(true); }
              }}>
                <Text style={styles.statusBoardHeaderText}>NAME</Text>
                {sortMode === 'name' && (
                  <View style={styles.sortBadge}>
                    <Text style={styles.sortBadgeText}>{sortAsc ? 'A-Z' : 'Z-A'}</Text>
                  </View>
                )}
              </Pressable>
              <Pressable style={[styles.statusBoardHeaderBtn, { flex: 1 }]} onPress={() => {
                if (sortMode === 'status') setSortAsc(!sortAsc);
                else { setSortMode('status'); setSortAsc(true); }
              }}>
                <Text style={styles.statusBoardHeaderText}>STATUS</Text>
                {sortMode === 'status' && (
                  <View style={styles.sortBadge}>
                    <Text style={styles.sortBadgeText}>{sortAsc ? 'WARN-UR' : 'UR-WARN'}</Text>
                  </View>
                )}
              </Pressable>
            </View>

            {filteredGroups.length === 0 ? (
              <View style={styles.statusBoardEmpty}>
                <Text style={styles.statusBoardEmptyText}>No patients found</Text>
              </View>
            ) : (
              filteredGroups.map((group) => {
                const latestRecord = group.records[group.records.length - 1];
                return (
                  <Pressable 
                    key={group.name.toLowerCase().trim()} 
                    style={({ pressed }) => [
                      styles.statusBoardRow,
                      pressed && styles.statusBoardRowPressed
                    ]}
                    onPress={() => {
                      onPatientPress(group.name);
                    }}
                  >
                    <Text style={[styles.statusBoardCell, { flex: 1.5 }]} numberOfLines={1}>
                      {group.name}
                    </Text>
                    <View style={{ flex: 1, alignItems: 'flex-start' }}>
                      {latestRecord.isUr ? (
                        <View style={styles.statusBoardBadgeOk}>
                          <Text style={styles.statusBoardBadgeTextOk}>✓ UR</Text>
                        </View>
                      ) : (
                        <View style={styles.statusBoardBadgeWarn}>
                          <Text style={styles.statusBoardBadgeTextWarn}>⚠ {latestRecord.breastValue || 'Suspicious'}</Text>
                        </View>
                      )}
                    </View>
                  </Pressable>
                );
              })
            )}
          </View>
        </ScrollView>
      </SafeAreaView>
    </Modal>
  );
}

function ImportExcelModal({
  visible,
  onClose,
  onImport,
}: {
  visible: boolean;
  onClose: () => void;
  onImport: (base64: string, fileName: string) => void;
}) {
  const [loading, setLoading] = useState(false);

  const handlePick = async () => {
    try {
      setLoading(true);
      const result = await DocumentPicker.getDocumentAsync({
        type: [
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'application/vnd.ms-excel',
        ],
        copyToCacheDirectory: true,
      });
      if (result.canceled || !result.assets?.[0]) {
        setLoading(false);
        return;
      }
      const asset = result.assets[0];
      const base64 = await FileSystem.readAsStringAsync(asset.uri, {
        encoding: FileSystem.EncodingType.Base64,
      });
      onImport(base64, asset.name);
      onClose();
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Failed to read file.';
      Alert.alert('Import failed', msg);
    } finally {
      setLoading(false);
    }
  };

  const downloadTemplate = async () => {
    try {
      const headers = [
        'Screening Date', 'ID No. ', 'Name', 'Age', 'Address', 'Contact No. ',
        'Civil Status ', 'No. of Children', 'Menarche', 'LMP', 'AOG',
        'Menstrual Bleeding Pattern ', 'No. of pads/day', 'Pregnancy History ',
        'Age at first full term pregnancy ', 'Oral Contraceptives Use',
        'Duration of usage of oral contraceptives',
        'History of previous cervical cancer screening ',
        'History of abnormal vaginal discharge', 'History of abnormal vaginal bleeding',
        'Age of first intercourse', 'No. of sexual partner', 'Spouse/Partner/s',
        'Family History of Cancer', 'Smoking ', 'Current medication', 'Allergies',
        'Abdominal surgery', 'BP', 'Temp', 'HR', 'RR', 'Height (cm)', 'Weight (kg)', 'BMI',
        'Skin', 'HEENT', 'Chest and Lungs', 'Heart', 'Abdomen', 'Right Breast', 'Left Breast',
      ];
      const sampleRow = [
        '06/28/2026', '001', 'Juan Dela Cruz', '45', 'Manila', '09123456789',
        'Married', '3', '13', '06/01/2026', '', 'Regular', '3',
        'G3T3P0A0L3', '22', 'Yes', '5 years', 'No', 'No', 'No',
        '20', '1', 'Circumcised', 'No', 'No', 'None', 'None', 'None',
        '120/80', '36.5', '80', '18', '160', '55', '21.5',
        'UR', 'UR', 'UR', 'UR', 'UR', 'UR', 'UR',
      ];
      const ws = XLSX.utils.aoa_to_sheet([headers, sampleRow]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Patient Template');
      const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
      const uri = `${FileSystem.cacheDirectory}team_silayan_template.xlsx`;
      await FileSystem.writeAsStringAsync(uri, base64, { encoding: FileSystem.EncodingType.Base64 });
      const canShare = await Sharing.isAvailableAsync();
      if (canShare) {
        await Sharing.shareAsync(uri, {
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          dialogTitle: 'Team Silayan - Patient Import Template',
        });
      } else {
        Alert.alert('Template saved', `File saved to cache: ${uri}`);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Failed to create template.';
      Alert.alert('Template failed', msg);
    }
  };

  return (
    <Modal visible={visible} animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <View style={styles.modalHeaderRow}>
          <Text style={styles.modalHeaderTitle}>Import Excel</Text>
          <Pressable style={styles.modalCancelButton} onPress={onClose}>
            <Text style={styles.modalCancelButtonText}>Cancel</Text>
          </Pressable>
        </View>
        <View style={{ flex: 1, padding: 24, gap: 16 }}>
          <Text style={{ color: '#4b5563', fontSize: 15, lineHeight: 22 }}>
            Select an Excel file (.xlsx) to import patient records. The file should have columns for{' '}
            <Text style={{ fontWeight: '700' }}>Name</Text>,{' '}
            <Text style={{ fontWeight: '700' }}>Screening Date</Text>,{' '}
            <Text style={{ fontWeight: '700' }}>Right Breast</Text>, and{' '}
            <Text style={{ fontWeight: '700' }}>Left Breast</Text>.
          </Text>
          <Text style={{ color: '#6b7280', fontSize: 13, lineHeight: 20 }}>
            New records will be merged with your existing data.
          </Text>

          <View style={{ marginTop: 8, gap: 12 }}>
            <Pressable
              style={[styles.saveBtn, loading && { opacity: 0.6 }]}
              onPress={handlePick}
              disabled={loading}
            >
              <LinearGradient
                colors={['#db4278', '#b51f55']}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 0 }}
                style={StyleSheet.absoluteFillObject}
              />
              <Text style={styles.saveBtnText}>{loading ? 'Loading...' : '📂 Choose Excel File'}</Text>
            </Pressable>

            <Pressable
              style={[styles.secondaryButton, { minHeight: 48 }]}
              onPress={downloadTemplate}
            >
              <Text style={styles.secondaryButtonText}>📥 Download Template</Text>
            </Pressable>
          </View>

          <View style={{ marginTop: 16, padding: 16, backgroundColor: '#f9fafb', borderRadius: 12, borderWidth: 1, borderColor: '#e5e7eb' }}>
            <Text style={{ color: '#374151', fontSize: 14, fontWeight: '700', marginBottom: 8 }}>Template Columns:</Text>
            <Text style={{ color: '#6b7280', fontSize: 12, lineHeight: 20 }}>
              Screening Date, ID No., Name, Age, Address, Contact No., Civil Status, No. of Children, Menarche, LMP, AOG, Menstrual Bleeding Pattern, No. of pads/day, Pregnancy History, Age at first full term pregnancy, Oral Contraceptives Use, Duration of usage of oral contraceptives, History of previous cervical cancer screening, History of abnormal vaginal discharge, History of abnormal vaginal bleeding, Age of first intercourse, No. of sexual partner, Spouse/Partner/s, Family History of Cancer, Smoking, Current medication, Allergies, Abdominal surgery, BP, Temp, HR, RR, Height (cm), Weight (kg), BMI, Skin, HEENT, Chest and Lungs, Heart, Abdomen, Right Breast, Left Breast
            </Text>
          </View>
        </View>
      </SafeAreaView>
    </Modal>
  );
}

const styles = StyleSheet.create({
  safeArea: {
    flex: 1,
    backgroundColor: '#faf8f9',
  },
  container: {
    flex: 1,
    paddingHorizontal: 16,
    paddingTop: 12,
    backgroundColor: '#faf8f9',
  },
  header: {
    marginBottom: 20,
    marginTop: Platform.OS === 'ios' ? 10 : 36,
    padding: 20,
    borderRadius: 16,
    borderWidth: 1,
    borderColor: '#f5d6e2',
    overflow: 'hidden',
    position: 'relative',
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 4 },
    shadowOpacity: 0.1,
    shadowRadius: 12,
    elevation: 3,
  },
  headerTop: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    zIndex: 1,
  },
  headerLogo: {
    height: 48,
    width: 100,
  },
  headerLogoRow: {
    flexDirection: 'row',
    alignItems: 'center',
    flex: 1,
    gap: 8,
  },
  headerIcon: {
    width: 56,
    height: 56,
  },
  title: {
    color: '#ffffff',
    fontSize: 28,
    fontWeight: '700',
    flex: 1,
  },
  aboutButton: {
    borderRadius: 999,
    paddingHorizontal: 14,
    paddingVertical: 8,
    backgroundColor: '#ffffff',
    borderWidth: 1,
    borderColor: '#ffffff',
  },
  aboutButtonText: {
    color: '#db4278',
    fontSize: 14,
    fontWeight: '600',
  },
  subtitle: {
    color: '#ffe4e6',
    fontSize: 15,
    lineHeight: 22,
    zIndex: 1,
    padding: 8,
    marginHorizontal: -20,
    marginBottom: -20,
    paddingHorizontal: 20,
  },
  button: {
    borderRadius: 12,
    paddingVertical: 14,
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: 48,
    marginBottom: 16,
    overflow: 'hidden',
    position: 'relative',
  },
  primaryButton: {
    backgroundColor: '#db4278',
  },
  buttonDisabled: {
    opacity: 0.7,
  },
  primaryButtonText: {
    color: '#ffffff',
    fontSize: 16,
    fontWeight: '600',
  },
  resultScroll: {
    flex: 1,
    zIndex: 1,
  },
  resultSection: {
    gap: 12,
    paddingBottom: 24,
  },
  fileCard: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    padding: 16,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
  },
  fileName: {
    color: '#1f2937',
    fontSize: 16,
    fontWeight: '600',
    marginBottom: 4,
  },
  fileMeta: {
    color: '#6b7280',
    fontSize: 14,
  },
  searchSection: {
    gap: 8,
    marginBottom: 12,
  },
  dateSearchRow: {
    flexDirection: 'row',
    gap: 12,
  },
  dateSearchGroup: {
    flex: 1,
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: '#ffffff',
    borderRadius: 10,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    paddingHorizontal: 10,
    paddingVertical: 4,
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 1 },
    shadowOpacity: 0.05,
    shadowRadius: 4,
    elevation: 1,
  },
  dateSearchLabel: {
    color: '#6b7280',
    fontSize: 12,
    fontWeight: '600',
    marginRight: 6,
  },
  dateSearchInput: {
    flex: 1,
    paddingVertical: 8,
    fontSize: 13,
    color: '#1f2937',
    minWidth: 80,
  },
  searchClearBtnSmall: {
    padding: 4,
  },
  searchFilterChip: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    backgroundColor: '#fff1f5',
    borderRadius: 8,
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderWidth: 1,
    borderColor: '#db4278',
  },
  searchFilterChipText: {
    color: '#db4278',
    fontSize: 12,
    fontWeight: '500',
  },
  searchFilterChipClear: {
    color: '#db4278',
    fontSize: 12,
    fontWeight: '700',
  },
  quickDateFilterRow: {
    flexDirection: 'row',
    gap: 8,
    flexWrap: 'wrap',
  },
  quickDateFilterBtn: {
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 6,
    backgroundColor: '#f3f4f6',
    borderWidth: 1,
    borderColor: '#e5e7eb',
  },
  quickDateFilterBtnActive: {
    backgroundColor: '#db4278',
    borderColor: '#db4278',
  },
  quickDateFilterBtnText: {
    color: '#6b7280',
    fontSize: 12,
    fontWeight: '600',
  },
  quickDateFilterBtnTextActive: {
    color: '#ffffff',
  },
  searchResultsInfo: {
    paddingHorizontal: 12,
    paddingVertical: 8,
    backgroundColor: '#f9fafb',
    borderBottomWidth: 1,
    borderBottomColor: '#f0e2e9',
  },
  searchResultsInfoText: {
    color: '#6b7280',
    fontSize: 12,
  },
  listCard: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    overflow: 'hidden',
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
  },
  listHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: 8,
    paddingVertical: 8,
    backgroundColor: '#ffe4e6',
    gap: 4,
  },
  listHeaderText: {
    color: '#9f1239',
    fontSize: 10,
    fontWeight: '700',
    letterSpacing: 0.3,
  },
  sortBadge: {
    backgroundColor: '#db4278',
    borderRadius: 3,
    paddingHorizontal: 4,
    paddingVertical: 1,
  },
  sortBadgeText: {
    color: '#ffffff',
    fontSize: 7,
    fontWeight: '700',
  },
  listRow: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: 8,
    paddingVertical: 10,
    borderBottomWidth: 1,
    borderBottomColor: '#f3e8ee',
    gap: 4,
    borderLeftWidth: 3,
    borderLeftColor: 'transparent',
  },
  listRowSelected: {
    backgroundColor: 'rgba(219, 66, 120, 0.08)',
    borderLeftColor: '#db4278',
  },
  listCell: {
    color: '#374151',
    fontSize: 12,
  },
  listNameCol: {
    flex: 1.8,
    minWidth: 60,
  },
  listNameText: {
    color: '#db4278',
    fontWeight: '600',
    textDecorationLine: 'underline',
    fontSize: 12,
  },
  listRecordCount: {
    color: '#9ca3af',
    fontSize: 9,
    marginTop: 1,
  },
  listStatusCol: {
    flex: 1.0,
    alignItems: 'center',
    justifyContent: 'center',
    minWidth: 70,
  },
  listScreeningCol: {
    flex: 1.2,
    minWidth: 50,
  },
  listReferralCol: {
    flex: 0.8,
    minWidth: 30,
  },
  statusBadgeWrap: {
    justifyContent: 'center',
    alignItems: 'center',
  },
  statusBadge: {
    borderRadius: 999,
    paddingHorizontal: 8,
    paddingVertical: 3,
    minWidth: 60,
    alignItems: 'center',
    justifyContent: 'center',
  },
  statusBadgeOk: {
    backgroundColor: '#d1fae5',
  },
  statusBadgeWarn: {
    backgroundColor: '#fef3c7',
  },
  statusBadgeText: {
    fontSize: 11,
    fontWeight: '700',
  },
  detailSection: {
    gap: 12,
  },
  detailSectionTitle: {
    color: '#1f2937',
    fontSize: 20,
    fontWeight: '700',
  },
  detailSectionMeta: {
    color: '#6b7280',
    fontSize: 14,
    marginTop: -4,
  },
  statusCard: {
    borderRadius: 16,
    padding: 28,
    alignItems: 'center',
    borderWidth: 1,
    overflow: 'hidden',
  },
  statusCardOk: {
    borderColor: '#a7f3d0',
  },
  statusCardWarn: {
    borderColor: '#fde68a',
  },
  statusIcon: {
    fontSize: 64,
    marginBottom: 12,
    zIndex: 1,
  },
  statusTitle: {
    fontSize: 22,
    fontWeight: '700',
    marginBottom: 8,
    zIndex: 1,
  },
  statusMessage: {
    fontSize: 15,
    textAlign: 'center',
    lineHeight: 22,
    zIndex: 1,
  },
  detailCard: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    padding: 16,
    gap: 12,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
  },
  detailCardTitle: {
    color: '#db4278',
    fontSize: 16,
    fontWeight: '700',
    marginBottom: 12,
  },
  categoryLabel: {
    color: '#6b7280',
    fontSize: 13,
    fontWeight: '700',
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginTop: 16,
    marginBottom: 4,
    borderBottomWidth: 1,
    borderBottomColor: '#f0e2e9',
    paddingBottom: 6,
  },
  categorySubLabel: {
    color: '#b51f55',
    fontSize: 13,
    fontWeight: '700',
    textTransform: 'uppercase',
    letterSpacing: 0.4,
    marginTop: 12,
    marginBottom: 4,
    paddingBottom: 4,
    borderBottomWidth: 1,
    borderBottomColor: '#fce7f0',
  },
  detailRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: 12,
  },
  detailLabel: {
    color: '#6b7280',
    fontSize: 14,
  },
  detailValue: {
    color: '#1f2937',
    fontSize: 15,
    fontWeight: '600',
    flexShrink: 1,
    textAlign: 'right',
  },
  detailValueHighlight: {
    color: '#db4278',
  },
  placeholderCard: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    padding: 16,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
    zIndex: 1,
  },
  placeholderTitle: {
    color: '#db4278',
    fontSize: 16,
    fontWeight: '600',
    marginBottom: 8,
  },
  placeholderText: {
    color: '#4b5563',
    fontSize: 14,
    lineHeight: 24,
  },
  aboutOverlay: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 24,
  },
  aboutBackdrop: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(0, 0, 0, 0.4)',
  },
  aboutCard: {
    width: '100%',
    maxWidth: 320,
    backgroundColor: '#ffffff',
    borderRadius: 16,
    padding: 24,
    alignItems: 'center',
    gap: 16,
    zIndex: 1,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    shadowColor: '#000000',
    shadowOffset: { width: 0, height: 4 },
    shadowOpacity: 0.1,
    shadowRadius: 12,
    elevation: 5,
  },
  aboutLogo: {
    width: 180,
    height: 180,
  },
  aboutDescription: {
    color: '#4b5563',
    fontSize: 14,
    lineHeight: 22,
    textAlign: 'center',
  },
  aboutCredit: {
    color: '#374151',
    fontSize: 16,
    fontWeight: '600',
    textAlign: 'center',
  },
  aboutCloseButton: {
    marginTop: 4,
    borderRadius: 10,
    paddingHorizontal: 20,
    paddingVertical: 10,
    backgroundColor: '#db4278',
    overflow: 'hidden',
  },
  aboutCloseButtonText: {
    color: '#ffffff',
    fontSize: 15,
    fontWeight: '600',
    zIndex: 1,
  },
  buttonRow: {
    flexDirection: 'row',
    gap: 12,
    marginBottom: 16,
    zIndex: 1,
  },
  rowButton: {
    flex: 1,
    marginBottom: 0,
  },
  secondaryButton: {
    backgroundColor: '#ffffff',
    borderWidth: 1,
    borderColor: '#db4278',
  },
  secondaryButtonText: {
    color: '#db4278',
    fontSize: 16,
    fontWeight: '600',
  },
  modalContainer: {
    flex: 1,
    backgroundColor: '#faf8f9',
  },
  modalHeaderRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: 20,
    paddingVertical: 16,
    borderBottomWidth: 1,
    borderBottomColor: '#f0e2e9',
    backgroundColor: '#ffffff',
    marginTop: Platform.OS === 'android' ? 24 : 0,
  },
  modalHeaderTitle: {
    color: '#1f2937',
    fontSize: 20,
    fontWeight: '700',
  },
  modalCancelButton: {
    paddingVertical: 6,
    paddingHorizontal: 12,
    borderRadius: 6,
    backgroundColor: '#e5e7eb',
  },
  modalCancelButtonText: {
    color: '#4b5563',
    fontSize: 14,
    fontWeight: '600',
  },
  modalScrollView: {
    flex: 1,
  },
  modalScrollContent: {
    padding: 20,
    paddingBottom: 40,
    gap: 16,
  },
  formGroup: {
    gap: 8,
  },
  formLabel: {
    color: '#4b5563',
    fontSize: 14,
    fontWeight: '600',
  },
  formInput: {
    backgroundColor: '#ffffff',
    borderWidth: 1,
    borderColor: '#d1d5db',
    borderRadius: 8,
    paddingHorizontal: 12,
    paddingVertical: 10,
    color: '#1f2937',
    fontSize: 15,
  },
  quickStatusRow: {
    flexDirection: 'row',
    gap: 12,
    marginTop: 4,
  },
  quickStatusBtn: {
    flex: 1,
    paddingVertical: 8,
    borderRadius: 6,
    alignItems: 'center',
    backgroundColor: '#f3f4f6',
    borderWidth: 1,
    borderColor: '#e5e7eb',
  },
  detailValueUr: {
    color: '#10b981', // Green for UR
    fontWeight: '700',
  },
  detailValueWarning: {
    color: '#d97706', // Amber/warning for anything else
    fontWeight: '700',
  },
  viewRecordFieldValueUr: {
    color: '#10b981', // Green for UR
    fontWeight: '700',
  },
  viewRecordFieldValueWarning: {
    color: '#d97706', // Amber/warning for anything else
    fontWeight: '700',
  },
  detailValueSuspicious: {
    color: '#d97706', // Amber/warning color for Suspicious
    fontWeight: '700',
  },  
  viewRecordFieldValueSuspicious: {
    color: '#d97706', // Amber/warning color for Suspicious
    fontWeight: '700',
  },
  quickStatusBtnActiveUr: {
    backgroundColor: '#d1fae5',
    borderColor: '#10b981',
  },
  quickStatusBtnActiveNotUr: {
    backgroundColor: '#fef3c7',
    borderColor: '#f59e0b',
  },
  quickStatusBtnText: {
    color: '#374151',
    fontSize: 13,
    fontWeight: '700',
  },
  sectionHeaderRow: {
    marginTop: 12,
    borderBottomWidth: 1,
    borderBottomColor: '#f0e2e9',
    paddingBottom: 6,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  sectionHeaderTitle: {
    color: '#db4278',
    fontSize: 16,
    fontWeight: '700',
  },
  customFieldRow: {
    flexDirection: 'row',
    gap: 8,
    alignItems: 'center',
  },
  customFieldLabelInput: {
    flex: 1.2,
  },
  customFieldValueInput: {
    flex: 2,
  },
  customFieldRemoveBtn: {
    padding: 10,
    borderRadius: 8,
    backgroundColor: '#fee2e2',
    borderWidth: 1,
    borderColor: '#fca5a5',
    width: 40,
    height: 40,
    alignItems: 'center',
    justifyContent: 'center',
  },
  customFieldRemoveText: {
    color: '#dc2626',
    fontSize: 14,
    fontWeight: '700',
  },
  addFieldBtn: {
    borderRadius: 8,
    paddingVertical: 10,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: '#d1d5db',
    borderStyle: 'dashed',
    marginTop: 8,
  },
  addFieldBtnText: {
    color: '#6b7280',
    fontSize: 14,
    fontWeight: '600',
  },
  saveBtn: {
    backgroundColor: '#db4278',
    borderRadius: 8,
    paddingVertical: 14,
    alignItems: 'center',
    justifyContent: 'center',
    marginTop: 20,
    overflow: 'hidden',
    position: 'relative',
  },
  saveBtnText: {
    color: '#ffffff',
    fontSize: 16,
    fontWeight: '600',
    zIndex: 1,
  },
  exportButton: {
    backgroundColor: '#ffffff',
    borderWidth: 1,
    borderColor: '#db4278',
  },
  exportButtonText: {
    color: '#db4278',
    fontSize: 16,
    fontWeight: '600',
  },
  admitNumberRow: {
    flexDirection: 'row',
    gap: 8,
    alignItems: 'center',
  },
  regenBtn: {
    paddingHorizontal: 12,
    paddingVertical: 10,
    borderRadius: 8,
    backgroundColor: '#ffffff',
    borderWidth: 1,
    borderColor: '#db4278',
  },
  regenBtnText: {
    color: '#db4278',
    fontSize: 14,
    fontWeight: '700',
  },
  addDateBtn: {
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 6,
    backgroundColor: '#fff1f5',
    borderWidth: 1,
    borderColor: '#db4278',
  },
  addDateBtnText: {
    color: '#db4278',
    fontSize: 13,
    fontWeight: '700',
  },
  referralDateRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
  },
  referralDateLabel: {
    color: '#6b7280',
    fontSize: 13,
    fontWeight: '600',
    width: 100,
  },
  referralDatesPanelCard: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    padding: 16,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    gap: 8,
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
  },
  referralDatesPanelHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: 4,
  },
  referralDatesPanelTitle: {
    color: '#1f2937',
    fontSize: 15,
    fontWeight: '700',
  },
  referralDatesPanelEntry: {
    color: '#374151',
    fontSize: 14,
    lineHeight: 22,
  },
  referralDatesPanelEmpty: {
    color: '#9ca3af',
    fontSize: 14,
    fontStyle: 'italic',
  },
  admissionHistoryCard: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    padding: 16,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    gap: 12,
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
  },
  admissionHistoryTitle: {
    color: '#1f2937',
    fontSize: 15,
    fontWeight: '700',
    marginBottom: 4,
  },
  admissionHistoryRow: {
    flexDirection: 'row',
    gap: 12,
    paddingVertical: 8,
    borderBottomWidth: 1,
    borderBottomColor: '#f3e8ee',
  },
  admissionHistoryNumber: {
    color: '#db4278',
    fontSize: 14,
    fontWeight: '700',
    width: 30,
  },
  admissionHistoryInfo: {
    flex: 1,
    gap: 2,
  },
  admissionHistoryDate: {
    color: '#374151',
    fontSize: 14,
    fontWeight: '600',
  },
  admissionHistoryAdmitNo: {
    color: '#6b7280',
    fontSize: 13,
  },
  admissionHistoryStatus: {
    color: '#6b7280',
    fontSize: 13,
  },
  admissionActionRow: {
    flexDirection: 'row',
    gap: 8,
    marginTop: 8,
  },
  viewRecordBtn: {
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 6,
    backgroundColor: '#eff6ff',
    borderWidth: 1,
    borderColor: '#3b82f6',
  },
  viewRecordBtnText: {
    color: '#2563eb',
    fontSize: 12,
    fontWeight: '700',
  },
  editRecordBtn: {
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 6,
    backgroundColor: '#fff1f5',
    borderWidth: 1,
    borderColor: '#db4278',
  },
  editRecordBtnText: {
    color: '#db4278',
    fontSize: 12,
    fontWeight: '700',
  },
  deleteRecordBtn: {
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 6,
    backgroundColor: '#fee2e2',
    borderWidth: 1,
    borderColor: '#fca5a5',
  },
  deleteRecordBtnText: {
    color: '#dc2626',
    fontSize: 12,
    fontWeight: '700',
  },
  editRecordPatientName: {
    color: '#1f2937',
    fontSize: 20,
    fontWeight: '700',
  },
  editRecordPatientMeta: {
    color: '#6b7280',
    fontSize: 14,
    marginTop: 2,
  },
  viewRecordHeader: {
    marginBottom: 16,
  },
  viewRecordName: {
    color: '#1f2937',
    fontSize: 22,
    fontWeight: '700',
  },
  viewRecordMeta: {
    color: '#6b7280',
    fontSize: 14,
    marginTop: 4,
  },
  viewRecordSection: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    padding: 16,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    marginBottom: 12,
    gap: 8,
  },
  viewRecordSectionTitle: {
    color: '#db4278',
    fontSize: 15,
    fontWeight: '700',
    marginBottom: 4,
  },
  viewRecordSubSectionTitle: {
    color: '#b51f55',
    fontSize: 13,
    fontWeight: '700',
    marginTop: 4,
    marginBottom: 2,
    borderBottomWidth: 1,
    borderBottomColor: '#fce7f0',
    paddingBottom: 2,
  },
  viewRecordItem: {
    color: '#374151',
    fontSize: 14,
    lineHeight: 22,
  },
  viewRecordEmpty: {
    color: '#9ca3af',
    fontSize: 14,
    fontStyle: 'italic',
  },
  viewRecordFieldRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    paddingVertical: 6,
    borderBottomWidth: 1,
    borderBottomColor: '#f3e8ee',
  },
  viewRecordFieldLabel: {
    color: '#6b7280',
    fontSize: 14,
  },
  viewRecordFieldValue: {
    color: '#1f2937',
    fontSize: 14,
    fontWeight: '600',
    flexShrink: 1,
    textAlign: 'right',
  },
  viewRecordFieldHighlight: {
    color: '#db4278',
  },
  exportSingleBtn: {
    borderRadius: 12,
    paddingVertical: 14,
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: 48,
    overflow: 'hidden',
    position: 'relative',
    marginTop: 8,
  },
  exportSingleBtnText: {
    color: '#ffffff',
    fontSize: 16,
    fontWeight: '600',
    zIndex: 1,
  },
  detailHeaderStack: {
    gap: 12,
  },
  detailNameContainer: {
    width: '100%',
  },
  addNewRecordBtn: {
    borderRadius: 10,
    paddingHorizontal: 12,
    paddingVertical: 10,
    overflow: 'hidden',
    position: 'relative',
    minWidth: 100,
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  addNewRecordBtnText: {
    color: '#ffffff',
    fontSize: 13,
    fontWeight: '700',
    zIndex: 1,
  },
  burgerButton: {
    width: 40,
    height: 40,
    justifyContent: 'center',
    alignItems: 'center',
    gap: 5,
    borderRadius: 8,
    backgroundColor: 'rgba(255,255,255,0.2)',
  },
  burgerLine: {
    width: 24,
    height: 2,
    backgroundColor: '#ffffff',
    borderRadius: 1,
  },
  sidebarOverlay: {
    flex: 1,
    flexDirection: 'row',
  },
  sidebarBackdrop: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(0, 0, 0, 0.4)',
  },
  sidebarCard: {
    width: 280,
    height: '100%',
    backgroundColor: '#ffffff',
    padding: 24,
    gap: 8,
    shadowColor: '#000000',
    shadowOffset: { width: 4, height: 0 },
    shadowOpacity: 0.2,
    shadowRadius: 12,
    elevation: 10,
  },
  sidebarHeader: {
    alignItems: 'center',
    gap: 12,
    marginBottom: 24,
    paddingBottom: 16,
    borderBottomWidth: 1,
    borderBottomColor: '#f0e2e9',
  },
  sidebarLogo: {
    width: 80,
    height: 80,
  },
  sidebarTitle: {
    color: '#db4278',
    fontSize: 20,
    fontWeight: '700',
  },
  sidebarItem: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 12,
    paddingVertical: 14,
    paddingHorizontal: 12,
    borderRadius: 10,
    backgroundColor: '#faf8f9',
    borderWidth: 1,
    borderColor: '#f0e2e9',
  },
  sidebarItemIconImage: {
    width: 28,
    height: 28,
    borderRadius: 6,
  },
  sidebarItemIconContainer: {
    width: 28,
    height: 28,
    borderRadius: 6,
    backgroundColor: '#ffffff',
    alignItems: 'center',
    justifyContent: 'center',
  },
  sidebarItemIconDanger: {
    fontSize: 14,
  },
  sidebarItemIconText: {
    fontSize: 14,
  },
  sidebarItemDanger: {
    backgroundColor: '#dc2626',
    borderColor: '#b91c1c',
    borderWidth: 1,
    borderRadius: 10,
    paddingVertical: 14,
    paddingHorizontal: 12,
    marginTop: 8,
    shadowColor: '#dc2626',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.2,
    shadowRadius: 4,
    elevation: 3,
  },
  sidebarItemTextDanger: {
    color: '#ffffff',
    fontSize: 16,
    fontWeight: '700',
  },
  staffAvatarImage: {
    width: 56,
    height: 56,
    borderRadius: 28,
    backgroundColor: '#f0e2e9',
  },
  staffFullscreenOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0, 0, 0, 0.92)',
    justifyContent: 'center',
    alignItems: 'center',
    padding: 20,
  },
  staffFullscreenImage: {
    width: '100%',
    height: '70%',
    borderRadius: 16,
  },
  staffFullscreenCloseBtn: {
    position: 'absolute',
    top: 50,
    right: 20,
    width: 44,
    height: 44,
    borderRadius: 22,
    backgroundColor: 'rgba(255, 255, 255, 0.2)',
    justifyContent: 'center',
    alignItems: 'center',
    zIndex: 10,
  },
  staffFullscreenCloseText: {
    color: '#ffffff',
    fontSize: 24,
    fontWeight: '700',
  },
  staffFullscreenName: {
    color: '#ffffff',
    fontSize: 20,
    fontWeight: '700',
    marginTop: 20,
    textAlign: 'center',
  },
  staffsIconContainer: {
    width: 28,
    height: 28,
    borderRadius: 6,
    backgroundColor: '#db4278',
    alignItems: 'center',
    justifyContent: 'center',
  },
  staffsIconText: {
    color: '#ffffff',
    fontSize: 12,
    fontWeight: '700',
  },
  sidebarItemText: {
    color: '#1f2937',
    fontSize: 16,
    fontWeight: '700',
  },
  sidebarCloseBtn: {
    marginTop: 'auto',
    borderRadius: 10,
    paddingHorizontal: 20,
    paddingVertical: 12,
    alignItems: 'center',
    overflow: 'hidden',
    position: 'relative',
  },
  sidebarCloseBtnText: {
    color: '#ffffff',
    fontSize: 15,
    fontWeight: '700',
    zIndex: 1,
  },
  staffsSubtitle: {
    color: '#6b7280',
    fontSize: 14,
    textAlign: 'center',
    marginBottom: 12,
  },
  staffCard: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 16,
    backgroundColor: '#ffffff',
    borderRadius: 12,
    padding: 16,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    marginBottom: 12,
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
  },
  staffAvatar: {
    width: 48,
    height: 48,
    borderRadius: 24,
    backgroundColor: '#db4278',
    alignItems: 'center',
    justifyContent: 'center',
  },
  staffAvatarText: {
    color: '#ffffff',
    fontSize: 18,
    fontWeight: '700',
  },
  staffInfo: {
    flex: 1,
    gap: 2,
  },
  staffName: {
    color: '#1f2937',
    fontSize: 16,
    fontWeight: '700',
  },
  staffRole: {
    color: '#db4278',
    fontSize: 14,
    fontWeight: '600',
  },
  staffDept: {
    color: '#6b7280',
    fontSize: 13,
  },
  backgroundLogo: {
    position: 'absolute',
    width: 300,
    height: 300,
    alignSelf: 'center',
    top: '30%',
    opacity: 0.08,
    zIndex: 0,
    pointerEvents: 'none',
  },
  // Carousel Overlay styles
  carouselOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0, 0, 0, 0.6)',
    justifyContent: 'center',
    alignItems: 'center',
  },
  carouselContainer: {
    backgroundColor: '#ffffff',
    borderRadius: 20,
    width: screenWidth - 32,
    maxHeight: screenHeight * 0.85,
    overflow: 'hidden',
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 4 },
    shadowOpacity: 0.3,
    shadowRadius: 12,
    elevation: 10,
  },
  carouselHeader: {
    alignItems: 'center',
    paddingTop: 20,
    paddingBottom: 12,
    borderBottomWidth: 1,
    borderBottomColor: '#f0e2e9',
  },
  carouselLogo: {
    width: 60,
    height: 60,
    marginBottom: 4,
  },
  carouselTitle: {
    color: '#db4278',
    fontSize: 20,
    fontWeight: '700',
  },
  carouselSubtitle: {
    color: '#6b7280',
    fontSize: 13,
    marginTop: 4,
    fontStyle: 'italic',
  },
  carouselFlatList: {
    height: 350,
  },
  carouselSlide: {
    width: screenWidth - 32,
    height: 350,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 12,
  },
  carouselImage: {
    width: '100%',
    height: '100%',
    borderRadius: 8,
  },
  carouselDots: {
    flexDirection: 'row',
    justifyContent: 'center',
    alignItems: 'center',
    paddingVertical: 12,
    gap: 8,
  },
  carouselDot: {
    width: 8,
    height: 8,
    borderRadius: 4,
    backgroundColor: '#d1d5db',
  },
  carouselDotActive: {
    backgroundColor: '#db4278',
    width: 20,
  },
  carouselCloseBtn: {
    marginHorizontal: 20,
    marginBottom: 20,
    borderRadius: 12,
    paddingVertical: 14,
    alignItems: 'center',
    justifyContent: 'center',
    overflow: 'hidden',
    position: 'relative',
  },
  carouselCloseBtnText: {
    color: '#ffffff',
    fontSize: 18,
    fontWeight: '700',
    zIndex: 1,
  },
  // Fullscreen Viewer styles
  fullscreenOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0, 0, 0, 0.95)',
    justifyContent: 'center',
    alignItems: 'center',
  },
  fullscreenHeader: {
    position: 'absolute',
    top: 40,
    left: 0,
    right: 0,
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    paddingHorizontal: 20,
    zIndex: 10,
  },
  fullscreenCounter: {
    color: '#ffffff',
    fontSize: 16,
    fontWeight: '600',
    backgroundColor: 'rgba(0, 0, 0, 0.5)',
    paddingHorizontal: 16,
    paddingVertical: 6,
    borderRadius: 20,
  },
  fullscreenCloseBtn: {
    width: 40,
    height: 40,
    borderRadius: 20,
    backgroundColor: 'rgba(255, 255, 255, 0.2)',
    justifyContent: 'center',
    alignItems: 'center',
  },
  fullscreenCloseBtnText: {
    color: '#ffffff',
    fontSize: 24,
    fontWeight: '700',
  },
  fullscreenFlatList: {
    flex: 1,
    width: screenWidth,
  },
  fullscreenSlide: {
    width: screenWidth,
    height: screenHeight,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 20,
  },
  fullscreenImage: {
    width: '100%',
    height: '90%',
  },
  fullscreenDots: {
    position: 'absolute',
    bottom: 40,
    flexDirection: 'row',
    justifyContent: 'center',
    alignItems: 'center',
    gap: 8,
  },
  fullscreenDot: {
    width: 8,
    height: 8,
    borderRadius: 4,
    backgroundColor: 'rgba(255, 255, 255, 0.3)',
  },
  fullscreenDotActive: {
    backgroundColor: '#db4278',
    width: 20,
  },
  // Status Board styles
  statusBoardListCard: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    overflow: 'hidden',
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
  },
  statusBoardListHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: 12,
    paddingVertical: 10,
    backgroundColor: '#ffe4e6',
    gap: 8,
  },
  statusBoardHeaderBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
  },
  statusBoardHeaderText: {
    color: '#9f1239',
    fontSize: 12,
    fontWeight: '700',
    letterSpacing: 0.5,
  },
  statusBoardRow: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: 12,
    paddingVertical: 14,
    borderBottomWidth: 1,
    borderBottomColor: '#f3e8ee',
    gap: 8,
    backgroundColor: '#ffffff',
  },
  statusBoardRowPressed: {
    backgroundColor: 'rgba(219, 66, 120, 0.05)',
  },
  statusBoardCell: {
    color: '#374151',
    fontSize: 14,
  },
  statusBoardBadgeOk: {
    backgroundColor: '#d1fae5',
    borderRadius: 999,
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderWidth: 1,
    borderColor: '#10b981',
  },
  statusBoardBadgeTextOk: {
    color: '#065f46',
    fontSize: 13,
    fontWeight: '700',
  },
  statusBoardBadgeWarn: {
    backgroundColor: '#fef3c7',
    borderRadius: 999,
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderWidth: 1,
    borderColor: '#f59e0b',
  },
  statusBoardBadgeTextWarn: {
    color: '#92400e',
    fontSize: 13,
    fontWeight: '700',
  },
  statusBoardEmpty: {
    padding: 40,
    alignItems: 'center',
    justifyContent: 'center',
  },
  statusBoardEmptyText: {
    color: '#9ca3af',
    fontSize: 16,
    fontStyle: 'italic',
  },
  // Grouped field boxes
  groupBox: {
    backgroundColor: '#ffffff',
    borderRadius: 12,
    padding: 16,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    gap: 12,
    marginTop: 8,
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.05,
    shadowRadius: 8,
    elevation: 2,
  },
  groupBoxTitle: {
    color: '#db4278',
    fontSize: 15,
    fontWeight: '700',
    marginBottom: 4,
  },
  listHeaderBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 3,
    paddingHorizontal: 2,
  },
  searchRow: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: '#ffffff',
    borderRadius: 10,
    borderWidth: 1,
    borderColor: '#f0e2e9',
    paddingHorizontal: 12,
    marginBottom: 8,
    shadowColor: '#db4278',
    shadowOffset: { width: 0, height: 1 },
    shadowOpacity: 0.05,
    shadowRadius: 4,
    elevation: 1,
  },
  searchInput: {
    flex: 1,
    paddingVertical: 10,
    fontSize: 14,
    color: '#1f2937',
  },
  searchClearBtn: {
    padding: 6,
  },
  searchClearBtnText: {
    color: '#9ca3af',
    fontSize: 14,
    fontWeight: '700',
  },
  admissionHistoryCardHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: 8,
  },
  historySortBtn: {
    paddingHorizontal: 10,
    paddingVertical: 5,
    borderRadius: 6,
    backgroundColor: '#fff1f5',
    borderWidth: 1,
    borderColor: '#db4278',
  },
  historySortBtnText: {
    color: '#db4278',
    fontSize: 12,
    fontWeight: '700',
  },
  historySearchRow: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: '#f9fafb',
    borderRadius: 8,
    borderWidth: 1,
    borderColor: '#e5e7eb',
    paddingHorizontal: 10,
    marginBottom: 10,
  },
  historySearchInput: {
    flex: 1,
    paddingVertical: 8,
    fontSize: 13,
    color: '#1f2937',
  },
  paginationRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    paddingVertical: 12,
    gap: 12,
    borderTopWidth: 1,
    borderTopColor: '#f3e8ee',
    marginTop: 4,
  },
  paginationBtn: {
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: 8,
    backgroundColor: '#fff1f5',
    borderWidth: 1,
    borderColor: '#db4278',
  },
  paginationBtnDisabled: {
    backgroundColor: '#f3f4f6',
    borderColor: '#d1d5db',
  },
  paginationBtnText: {
    color: '#db4278',
    fontSize: 13,
    fontWeight: '700',
  },
  paginationBtnTextDisabled: {
    color: '#9ca3af',
  },
  paginationText: {
    color: '#6b7280',
    fontSize: 13,
    fontWeight: '600',
  },
});

export default App;