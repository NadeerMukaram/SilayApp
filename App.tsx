import { useCallback, useState, useEffect, useMemo } from 'react';
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
} from 'react-native';
import { StatusBar } from 'expo-status-bar';
import { LinearGradient } from 'expo-linear-gradient';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';
import * as XLSX from 'xlsx';
import DateTimePicker from '@react-native-community/datetimepicker';

import { ExcelParseResult, PatientRecord } from './src/utils/excel';

const teamSilayanLogo = require('./assets/team-silayan.png');

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
};

const STORAGE_KEY = '@team_silayan_patient_records';

export default function App() {
  const [loadedFile, setLoadedFile] = useState<LoadedFile | null>(null);
  const [selectedPatientName, setSelectedPatientName] = useState<string | null>(null);
  const [aboutVisible, setAboutVisible] = useState(false);
  const [addPatientVisible, setAddPatientVisible] = useState(false);
  const [editRecordPatient, setEditRecordPatient] = useState<PatientRecord | null>(null);
  const [viewRecordPatient, setViewRecordPatient] = useState<PatientRecord | null>(null);
  const [addRecordForPatient, setAddRecordForPatient] = useState<string | null>(null);
  const [sidebarVisible, setSidebarVisible] = useState(false);
  const [staffsVisible, setStaffsVisible] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  // Load saved data on mount
  useEffect(() => {
    const loadData = async () => {
      try {
        const saved = await AsyncStorage.getItem(STORAGE_KEY);
        if (saved) {
          const parsed = JSON.parse(saved);
          if (parsed.name && parsed.result && parsed.result.patients) {
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

  const [sortMode, setSortMode] = useState<'name' | 'date' | 'status'>('name');
  const [sortAsc, setSortAsc] = useState(true);

  const sortedPatientGroups = useMemo(() => {
    const sorted = [...patientGroups];
    sorted.sort((a, b) => {
      let cmp = 0;
      if (sortMode === 'name') {
        cmp = a.name.localeCompare(b.name);
      } else if (sortMode === 'date') {
        const dateA = new Date(a.records[a.records.length - 1]?.date || '').getTime() || 0;
        const dateB = new Date(b.records[b.records.length - 1]?.date || '').getTime() || 0;
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
      breastValue: string;
      fields: Record<string, string>;
    }) => {
      const isUr = patientData.breastValue.toUpperCase() === 'UR';
      const maxRowIndex = loadedFile?.result.patients.reduce((max, p) => Math.max(max, p.rowIndex), 0) ?? 0;
      const nextRowIndex = maxRowIndex + 1;

      const newPatient: PatientRecord = {
        rowIndex: nextRowIndex,
        date: patientData.date,
        patientName: patientData.patientName,
        breastValue: patientData.breastValue || '(empty)',
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
    },
    [loadedFile],
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

  const exportExcel = useCallback(async () => {
    if (!loadedFile?.result.patients.length) {
      Alert.alert('Nothing to export', 'Create some patient records first.');
      return;
    }
    try {
      const patients = loadedFile.result.patients;
      // Get all unique field keys
      const allKeys = new Set<string>();
      patients.forEach((p) => Object.keys(p.fields).forEach((k) => allKeys.add(k)));
      const headers = Array.from(allKeys);

      // Expand patients with multiple dates into separate rows
      const rows: string[][] = [];
      patients.forEach((p) => {
        const dates = p.date && p.date !== '(no date)' && p.date !== '(empty)'
          ? p.date.split(';').map((d) => d.trim()).filter(Boolean)
          : [''];

        dates.forEach((date) => {
          const row = headers.map((h) => {
            if (h.toLowerCase().includes('date') || h.toLowerCase() === 'date') {
              return date;
            }
            return p.fields[h] ?? '';
          });
          rows.push(row);
        });
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
          <Text style={styles.subtitle}>
            Add patient records manually. Each patient is listed with name, latest date, and breast status.
            Tap a name to view full admission history and details.
          </Text>
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
                <Pressable style={[styles.listHeaderBtn, styles.listDateCol]} onPress={() => {
                  if (sortMode === 'date') setSortAsc(!sortAsc);
                  else { setSortMode('date'); setSortAsc(true); }
                }}>
                  <Text style={styles.listHeaderText}>LATEST DATE</Text>
                  {sortMode === 'date' && (
                    <View style={styles.sortBadge}>
                      <Text style={styles.sortBadgeText}>{sortAsc ? 'OLD-NEW' : 'NEW-OLD'}</Text>
                    </View>
                  )}
                </Pressable>
                <Pressable style={[styles.listHeaderBtn, styles.listStatusCol]} onPress={() => {
                  if (sortMode === 'status') setSortAsc(!sortAsc);
                  else { setSortMode('status'); setSortAsc(true); }
                }}>
                  <Text style={styles.listHeaderText}>BREAST</Text>
                  {sortMode === 'status' && (
                    <View style={styles.sortBadge}>
                      <Text style={styles.sortBadgeText}>{sortAsc ? 'WARN-UR' : 'UR-WARN'}</Text>
                    </View>
                  )}
                </Pressable>
              </View>

              {sortedPatientGroups.map((group) => (
                <PatientListRow
                  key={group.name.toLowerCase().trim()}
                  group={group}
                  selected={selectedPatientName?.toLowerCase().trim() === group.name.toLowerCase().trim()}
                  onPress={() => {
                    setSelectedPatientName((current) =>
                      current?.toLowerCase().trim() === group.name.toLowerCase().trim() ? null : group.name,
                    );
                  }}
                />
              ))}
            </View>

            {selectedPatientName && selectedRecords.length > 0 ? (
              <PatientDetail
                patientName={selectedPatientName}
                records={selectedRecords}
                onViewRecord={(p) => setViewRecordPatient(p)}
                onEditRecord={(p) => setEditRecordPatient(p)}
                onAddNewRecord={() => setAddRecordForPatient(selectedPatientName)}
                onDeleteRecord={(p) => {
                  Alert.alert(
                    'Delete Record',
                    `Are you sure you want to delete this record for ${p.patientName}?

Admit No: ${p.fields['Admit Number'] || 'N/A'}`,
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
                          if (updatedPatients.filter(
                            (pt) => pt.patientName.toLowerCase().trim() === p.patientName.toLowerCase().trim()
                          ).length === 0) {
                            setSelectedPatientName(null);
                          }
                        }
                      },
                    ]
                  );
                }}
              />
            ) : null}
          </ScrollView>
        ) : (
          <View style={styles.placeholderCard}>
            <Text style={styles.placeholderTitle}>How it works</Text>
            <Text style={styles.placeholderText}>1. Tap "Add Patient" to create a new patient record</Text>
            <Text style={styles.placeholderText}>2. All patients appear in the list</Text>
            <Text style={styles.placeholderText}>3. List shows name, latest date, and breast status</Text>
            <Text style={styles.placeholderText}>4. Tap a name to see full admission history</Text>
            <Text style={styles.placeholderText}>5. Tap "View" to see individual record details</Text>
            <Text style={styles.placeholderText}>6. Tap "Edit" to modify any record data</Text>
          </View>
        )}
      </View>

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
        onClearData={clearAllData}
        hasData={!!loadedFile}
      />
      <AboutModal visible={aboutVisible} onClose={() => setAboutVisible(false)} />
      <StaffsModal visible={staffsVisible} onClose={() => setStaffsVisible(false)} />
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
          onClose={() => setAddRecordForPatient(null)}
          onSave={(data) => {
            handleSavePatient(data);
            setAddRecordForPatient(null);
          }}
          existingPatients={loadedFile?.result.patients}
        />
      )}
    </SafeAreaView>
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
  const latestDate = latestRecord.date && latestRecord.date !== '(no date)' && latestRecord.date !== '(empty)'
    ? latestRecord.date.split(';')[0].trim()
    : '(no date)';

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
      <Text style={[styles.listCell, styles.listDateCol]} numberOfLines={1}>
        {latestDate}
      </Text>
      <View style={[styles.listStatusCol, styles.statusBadgeWrap]}>
        <View style={[styles.statusBadge, latestRecord.isUr ? styles.statusBadgeOk : styles.statusBadgeWarn]}>
          <Text style={[styles.statusBadgeText, { color: latestRecord.isUr ? '#065f46' : '#92400e' }]}>
            {latestRecord.isUr ? '✓' : '⚠'}
          </Text>
        </View>
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

      <View style={styles.admissionHistoryCard}>
        <Text style={styles.admissionHistoryTitle}>Admission History</Text>
        {records.map((record, idx) => {
          const admitNo = record.fields['Admit Number'] || record.fields['admit number'] || 'N/A';
          const recordDates = record.date && record.date !== '(no date)' && record.date !== '(empty)'
            ? record.date.split(';').map(d => d.trim()).filter(Boolean)
            : [];

          return (
            <View key={record.rowIndex} style={styles.admissionHistoryRow}>
              <Text style={styles.admissionHistoryNumber}>#{idx + 1}</Text>
              <View style={styles.admissionHistoryInfo}>
                <Text style={styles.admissionHistoryAdmitNo}>Admit No: {admitNo}</Text>
                {recordDates.length > 0 && (
                  <Text style={styles.admissionHistoryDate}>
                    Referral Date{recordDates.length > 1 ? 's' : ''}: {recordDates.join(', ')}
                  </Text>
                )}
                <Text style={styles.admissionHistoryStatus}>
                  Breast: {record.isUr ? 'UR' : record.breastValue}
                </Text>
                <View style={styles.admissionActionRow}>
                  <Pressable
                    style={styles.viewRecordBtn}
                    onPress={() => onViewRecord(record)}
                  >
                    <Text style={styles.viewRecordBtnText}>View</Text>
                  </Pressable>
                  <Pressable
                    style={styles.editRecordBtn}
                    onPress={() => onEditRecord(record)}
                  >
                    <Text style={styles.editRecordBtnText}>Edit</Text>
                  </Pressable>
                  <Pressable
                    style={styles.deleteRecordBtn}
                    onPress={() => onDeleteRecord(record)}
                  >
                    <Text style={styles.deleteRecordBtnText}>Delete</Text>
                  </Pressable>
                </View>
              </View>
            </View>
          );
        })}
      </View>

      <View style={styles.detailCard}>
        <Text style={styles.detailCardTitle}>Latest Record Details</Text>
        {Object.entries(records[records.length - 1].fields).map(([label, value]) => (
          <DetailRow
            key={label}
            label={label}
            value={value}
            highlight={label.toLowerCase().includes('breast')}
          />
        ))}
      </View>
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
  return (
    <View style={styles.detailRow}>
      <Text style={styles.detailLabel}>{label.replace(/\b\w/g, (c) => c.toUpperCase())}</Text>
      <Text style={[styles.detailValue, highlight && styles.detailValueHighlight]}>{value}</Text>
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
  // Referral dates removed from view - shown in admission history instead

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
            <Text style={styles.viewRecordMeta}>Row {patient.rowIndex} · Admit No: {patient.fields['Admit Number'] || 'N/A'}</Text>
          </View>

          <View
            style={[
              styles.statusCard,
              patient.isUr ? styles.statusCardOk : styles.statusCardWarn,
              { marginBottom: 16 },
            ]}
          >
            <LinearGradient
              colors={patient.isUr ? ['#eafaf1', '#d1f2e1'] : ['#fdf6e2', '#f9e8be']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 1 }}
              style={StyleSheet.absoluteFillObject}
            />
            <Text style={[styles.statusIcon, { color: patient.isUr ? '#10b981' : '#f59e0b', fontSize: 48 }]}>
              {patient.isUr ? '✓' : '⚠'}
            </Text>
            <Text style={[styles.statusTitle, { color: patient.isUr ? '#065f46' : '#92400e' }]}>
              {patient.isUr ? 'Breast: UR' : 'Breast: Warning'}
            </Text>
            <Text style={[styles.statusMessage, { color: patient.isUr ? '#0f5132' : '#664d03' }]}>
              {patient.isUr
                ? 'The Breast value is Unremarkable.'
                : `The Breast value is "${patient.breastValue}".`}
            </Text>
          </View>

          <View style={styles.viewRecordSection}>
            <Text style={styles.viewRecordSectionTitle}>All Fields</Text>
            {Object.entries(patient.fields).map(([label, value]) => (
              <View key={label} style={styles.viewRecordFieldRow}>
                <Text style={styles.viewRecordFieldLabel}>{label.replace(/\b\w/g, (c) => c.toUpperCase())}</Text>
                <Text style={[styles.viewRecordFieldValue, label.toLowerCase().includes('breast') && styles.viewRecordFieldHighlight]}>
                  {value || '(empty)'}
                </Text>
              </View>
            ))}
          </View>

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
  const [name, setName] = useState(patient.patientName);
  const [referralDates, setReferralDates] = useState<Array<{ id: string; value: string }>>([]);
  const [breast, setBreast] = useState(patient.breastValue);
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const [datePicker, setDatePicker] = useState<DatePickerState>({
    visible: false,
    mode: 'single',
    targetId: null,
    currentValue: '',
  });

  useEffect(() => {
    setName(patient.patientName);
    setBreast(patient.breastValue);

    const raw = patient.date || '';
    if (raw && raw !== '(no date)' && raw !== '(empty)') {
      const dates = raw.split(';').map((d) => d.trim()).filter(Boolean);
      setReferralDates(dates.map((d, i) => ({ id: `existing-${i}`, value: d })));
    } else {
      setReferralDates([{ id: 'init', value: todayString() }]);
    }

    const otherFields: Record<string, string> = {};
    Object.entries(patient.fields).forEach(([key, value]) => {
      const lower = key.toLowerCase();
      if (!lower.includes('name') && !lower.includes('date') && !lower.includes('breast')) {
        otherFields[key] = value;
      }
    });
    setFieldValues(otherFields);
  }, [patient]);

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

  const openDatePicker = (id: string, currentValue: string) => {
    setDatePicker({
      visible: true,
      mode: 'edit',
      targetId: id,
      currentValue,
    });
  };

  const onDateChange = (event: any, selectedDate?: Date) => {
    if (selectedDate) {
      const mm = String(selectedDate.getMonth() + 1).padStart(2, '0');
      const dd = String(selectedDate.getDate()).padStart(2, '0');
      const yyyy = selectedDate.getFullYear();
      const formatted = `${mm}/${dd}/${yyyy}`;

      if (datePicker.mode === 'edit' && datePicker.targetId) {
        updateReferralDate(datePicker.targetId, formatted);
      }
    }
    setDatePicker((prev) => ({ ...prev, visible: false }));
  };

  const handleSave = () => {
    if (!name.trim()) {
      Alert.alert('Validation Error', 'Patient Name is required.');
      return;
    }

    const combinedDates = referralDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryDate = referralDates[0]?.value.trim() || '(no date)';

    const isUr = breast.toUpperCase() === 'UR';

    const updatedFields: Record<string, string> = { ...fieldValues };

    const allKeys = Object.keys(patient.fields);
    const nameKey = allKeys.find((k) => k.toLowerCase() === 'name' || k.toLowerCase().includes('patient name')) || 'name';
    const dateKey = allKeys.find((k) => k.toLowerCase() === 'date' || k.toLowerCase().includes('referral date')) || 'date';
    const breastKey = allKeys.find((k) => k.toLowerCase() === 'breast') || 'Breast';

    updatedFields[nameKey] = name.trim();
    updatedFields[dateKey] = combinedDates;
    updatedFields[breastKey] = breast.trim();

    const admitNoKey = allKeys.find((k) => k.toLowerCase().includes('admit number'));
    if (admitNoKey) {
      updatedFields[admitNoKey] = patient.fields[admitNoKey];
    }

    const updatedPatient: PatientRecord = {
      ...patient,
      patientName: name.trim(),
      date: primaryDate,
      breastValue: breast.trim(),
      isUr,
      fields: updatedFields,
    };

    onSave(updatedPatient);
  };

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
              <Text style={styles.sectionHeaderTitle}>Referral Dates</Text>
              <Pressable style={styles.addDateBtn} onPress={addReferralDate}>
                <Text style={styles.addDateBtnText}>+ Add Date</Text>
              </Pressable>
            </View>

            {referralDates.map((d, idx) => (
              <View style={styles.referralDateRow} key={d.id}>
                <Text style={styles.referralDateLabel}>Date {idx + 1}</Text>
                <Pressable
                  style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                  onPress={() => openDatePicker(d.id, d.value)}
                >
                  <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                </Pressable>
                {referralDates.length > 1 && (
                  <Pressable
                    style={styles.customFieldRemoveBtn}
                    onPress={() => removeReferralDate(d.id)}
                  >
                    <Text style={styles.customFieldRemoveText}>X</Text>
                  </Pressable>
                )}
              </View>
            ))}

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
              <Text style={styles.formLabel}>Breast</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Warning"
                placeholderTextColor="#64748b"
                value={breast}
                onChangeText={setBreast}
              />
              <View style={styles.quickStatusRow}>
                <Pressable
                  style={[
                    styles.quickStatusBtn,
                    breast.toUpperCase() === 'UR' && styles.quickStatusBtnActiveUr,
                  ]}
                  onPress={() => setBreast('UR')}
                >
                  <Text style={[
                    styles.quickStatusBtnText,
                    breast.toUpperCase() === 'UR' && { color: '#065f46' }
                  ]}>Set UR</Text>
                </Pressable>
                <Pressable
                  style={[
                    styles.quickStatusBtn,
                    breast.toUpperCase() !== 'UR' && breast !== '' && styles.quickStatusBtnActiveNotUr,
                  ]}
                  onPress={() => setBreast('Warning')}
                >
                  <Text style={[
                    styles.quickStatusBtnText,
                    breast.toUpperCase() !== 'UR' && breast !== '' && { color: '#92400e' }
                  ]}>Warning</Text>
                </Pressable>
              </View>
            </View>

            <View style={styles.sectionHeaderRow}>
              <Text style={styles.sectionHeaderTitle}>Other Fields</Text>
            </View>

            {Object.entries(fieldValues).map(([key, value]) => (
              <View style={styles.formGroup} key={key}>
                <Text style={styles.formLabel}>{key.replace(/\b\w/g, (c) => c.toUpperCase())}</Text>
                <TextInput
                  style={styles.formInput}
                  placeholder={`Enter ${key.toLowerCase()}`}
                  placeholderTextColor="#64748b"
                  value={value}
                  onChangeText={(text) =>
                    setFieldValues((prev) => ({ ...prev, [key]: text }))
                  }
                />
              </View>
            ))}

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

function checkDuplicateDate(
  name: string,
  newDateValues: string[],
  existingPatients: PatientRecord[] | undefined,
  onProceed: () => void,
): boolean {
  if (!existingPatients || existingPatients.length === 0) {
    onProceed();
    return true;
  }

  const sameNamePatients = existingPatients.filter(
    (p) => p.patientName.toLowerCase().trim() === name.toLowerCase().trim()
  );

  for (const existing of sameNamePatients) {
    const existingDates = existing.date.split(';').map((d) => d.trim()).filter(Boolean);
    for (const newDate of newDateValues) {
      if (existingDates.includes(newDate)) {
        Alert.alert(
          'Duplicate Date Warning',
          `The referral date "${newDate}" already exists for patient "${existing.patientName}".

Do you want to continue?`,
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
    breastValue: string;
    fields: Record<string, string>;
  }) => void;
  existingPatients?: PatientRecord[];
}) {
  const [name, setName] = useState('');
  const [referralDates, setReferralDates] = useState<Array<{ id: string; value: string }>>([]);
  const [admitNumber, setAdmitNumber] = useState('');
  const [breast, setBreast] = useState('UR');
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const [customFields, setCustomFields] = useState<Array<{ id: string; label: string; value: string }>>([]);
  const [datePicker, setDatePicker] = useState<DatePickerState>({
    visible: false,
    mode: 'single',
    targetId: null,
    currentValue: '',
  });

  const fieldKeys = useMemo(() => {
    if (existingPatients && existingPatients.length > 0) {
      const keys = Object.keys(existingPatients[0].fields);
      return keys.filter((key) => {
        const lowerKey = key.toLowerCase();
        return (
          !lowerKey.includes('name') &&
          !lowerKey.includes('date') &&
          !lowerKey.includes('breast') &&
          !lowerKey.includes('admit number') &&
          !lowerKey.includes('admit no')
        );
      });
    }
    return [
      'ID no. (Hosp)',
      'age',
      'address',
      'civil status',
      'BP',
      'Temp',
      'PR',
      'RR',
      'family hx of cancer',
      'Smoking hx',
      'Allergies',
    ];
  }, [existingPatients]);

  useEffect(() => {
    if (visible) {
      setName('');
      setReferralDates([{ id: 'init', value: todayString() }]);
      setAdmitNumber(generateAdmitNumber());
      setBreast('UR');
      const initialValues: Record<string, string> = {};
      fieldKeys.forEach((key) => { initialValues[key] = ''; });
      setFieldValues(initialValues);
      setCustomFields([]);
    }
  }, [visible, fieldKeys]);

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

  const openDatePicker = (id: string, currentValue: string) => {
    setDatePicker({
      visible: true,
      mode: 'single',
      targetId: id,
      currentValue,
    });
  };

  const onDateChange = (event: any, selectedDate?: Date) => {
    if (selectedDate) {
      const mm = String(selectedDate.getMonth() + 1).padStart(2, '0');
      const dd = String(selectedDate.getDate()).padStart(2, '0');
      const yyyy = selectedDate.getFullYear();
      const formatted = `${mm}/${dd}/${yyyy}`;

      if (datePicker.targetId) {
        updateReferralDate(datePicker.targetId, formatted);
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
    const combinedDates = referralDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryDate = referralDates[0]?.value.trim() || '(no date)';

    const finalFields: Record<string, string> = {};

    let nameKey = 'name';
    let dateKey = 'date';
    let breastKey = 'Breast';

    if (existingPatients && existingPatients.length > 0) {
      const firstPatientKeys = Object.keys(existingPatients[0].fields);

      const foundNameKey = firstPatientKeys.find(
        (k) => k.toLowerCase() === 'name' || k.toLowerCase().includes('patient name'),
      );
      if (foundNameKey) nameKey = foundNameKey;

      const foundDateKey = firstPatientKeys.find((k) => k.toLowerCase() === 'date' || k.toLowerCase().includes('referral date'));
      if (foundDateKey) dateKey = foundDateKey;

      const foundBreastKey = firstPatientKeys.find((k) => k.toLowerCase() === 'breast');
      if (foundBreastKey) breastKey = foundBreastKey;

      firstPatientKeys.forEach((k) => {
        if (k === nameKey) {
          finalFields[k] = name.trim();
        } else if (k === dateKey) {
          finalFields[k] = combinedDates;
        } else if (k === breastKey) {
          finalFields[k] = breast.trim();
        } else {
          finalFields[k] = (fieldValues[k] || '').trim();
        }
      });
    } else {
      finalFields['Referral Date'] = combinedDates;
      finalFields['Admit Number'] = admitNumber.trim();
      finalFields['name'] = name.trim();
      finalFields['Breast'] = breast.trim();
      Object.entries(fieldValues).forEach(([k, v]) => {
        finalFields[k] = (v || '').trim();
      });
    }

    finalFields['Admit Number'] = admitNumber.trim();

    customFields.forEach((cf) => {
      if (cf.label.trim()) {
        finalFields[cf.label.trim()] = cf.value.trim();
      }
    });

    onSave({
      patientName: name.trim(),
      date: primaryDate,
      breastValue: breast.trim(),
      fields: finalFields,
    });
  };

  const handleSave = () => {
    if (!name.trim()) {
      Alert.alert('Validation Error', 'Patient Name is required.');
      return;
    }

    const newDateValues = referralDates.map((d) => d.value.trim()).filter(Boolean);
    checkDuplicateDate(name.trim(), newDateValues, existingPatients, proceedWithSave);
  };

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
              <Text style={styles.sectionHeaderTitle}>Referral Dates</Text>
              <Pressable style={styles.addDateBtn} onPress={addReferralDate}>
                <Text style={styles.addDateBtnText}>+ Add Date</Text>
              </Pressable>
            </View>

            {referralDates.map((d, idx) => (
              <View style={styles.referralDateRow} key={d.id}>
                <Text style={styles.referralDateLabel}>Referral Date {idx + 1}</Text>
                <Pressable
                  style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                  onPress={() => openDatePicker(d.id, d.value)}
                >
                  <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                </Pressable>
                {referralDates.length > 1 && (
                  <Pressable
                    style={styles.customFieldRemoveBtn}
                    onPress={() => removeReferralDate(d.id)}
                  >
                    <Text style={styles.customFieldRemoveText}>X</Text>
                  </Pressable>
                )}
              </View>
            ))}

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
              <Text style={styles.formLabel}>Breast</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Warning"
                placeholderTextColor="#64748b"
                value={breast}
                onChangeText={setBreast}
              />
              <View style={styles.quickStatusRow}>
                <Pressable
                  style={[
                    styles.quickStatusBtn,
                    breast.toUpperCase() === 'UR' && styles.quickStatusBtnActiveUr,
                  ]}
                  onPress={() => setBreast('UR')}
                >
                  <Text style={[
                    styles.quickStatusBtnText,
                    breast.toUpperCase() === 'UR' && { color: '#065f46' }
                  ]}>Set UR</Text>
                </Pressable>
                <Pressable
                  style={[
                    styles.quickStatusBtn,
                    breast.toUpperCase() !== 'UR' && breast !== '' && styles.quickStatusBtnActiveNotUr,
                  ]}
                  onPress={() => setBreast('Warning')}
                >
                  <Text style={[
                    styles.quickStatusBtnText,
                    breast.toUpperCase() !== 'UR' && breast !== '' && { color: '#92400e' }
                  ]}>Warning</Text>
                </Pressable>
              </View>
            </View>

            <View style={styles.sectionHeaderRow}>
              <Text style={styles.sectionHeaderTitle}>Additional Details</Text>
            </View>

            {fieldKeys.map((key) => (
              <View style={styles.formGroup} key={key}>
                <Text style={styles.formLabel}>{key.replace(/\b\w/g, (c) => c.toUpperCase())}</Text>
                <TextInput
                  style={styles.formInput}
                  placeholder={`Enter ${key.toLowerCase()}`}
                  placeholderTextColor="#64748b"
                  value={fieldValues[key]}
                  onChangeText={(text) =>
                    setFieldValues((prev) => ({ ...prev, [key]: text }))
                  }
                />
              </View>
            ))}

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
    breastValue: string;
    fields: Record<string, string>;
  }) => void;
  existingPatients?: PatientRecord[];
}) {
  const [name, setName] = useState(patientName);
  const [referralDates, setReferralDates] = useState<Array<{ id: string; value: string }>>([{ id: 'init', value: todayString() }]);
  const [admitNumber, setAdmitNumber] = useState(generateAdmitNumber());
  const [breast, setBreast] = useState('UR');
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const [customFields, setCustomFields] = useState<Array<{ id: string; label: string; value: string }>>([]);
  const [datePicker, setDatePicker] = useState<DatePickerState>({
    visible: false,
    mode: 'single',
    targetId: null,
    currentValue: '',
  });

  const fieldKeys = useMemo(() => {
    if (existingPatients && existingPatients.length > 0) {
      const keys = Object.keys(existingPatients[0].fields);
      return keys.filter((key) => {
        const lowerKey = key.toLowerCase();
        return (
          !lowerKey.includes('name') &&
          !lowerKey.includes('date') &&
          !lowerKey.includes('breast') &&
          !lowerKey.includes('admit number') &&
          !lowerKey.includes('admit no')
        );
      });
    }
    return [
      'ID no. (Hosp)',
      'age',
      'address',
      'civil status',
      'BP',
      'Temp',
      'PR',
      'RR',
      'family hx of cancer',
      'Smoking hx',
      'Allergies',
    ];
  }, [existingPatients]);

  useEffect(() => {
    setName(patientName);
    setReferralDates([{ id: 'init', value: todayString() }]);
    setAdmitNumber(generateAdmitNumber());
    setBreast('UR');
    const initialValues: Record<string, string> = {};
    fieldKeys.forEach((key) => { initialValues[key] = ''; });
    setFieldValues(initialValues);
    setCustomFields([]);
  }, [patientName, fieldKeys]);

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

  const openDatePicker = (id: string, currentValue: string) => {
    setDatePicker({
      visible: true,
      mode: 'single',
      targetId: id,
      currentValue,
    });
  };

  const onDateChange = (event: any, selectedDate?: Date) => {
    if (selectedDate) {
      const mm = String(selectedDate.getMonth() + 1).padStart(2, '0');
      const dd = String(selectedDate.getDate()).padStart(2, '0');
      const yyyy = selectedDate.getFullYear();
      const formatted = `${mm}/${dd}/${yyyy}`;

      if (datePicker.targetId) {
        updateReferralDate(datePicker.targetId, formatted);
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
    const combinedDates = referralDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryDate = referralDates[0]?.value.trim() || '(no date)';

    const finalFields: Record<string, string> = {};

    let nameKey = 'name';
    let dateKey = 'date';
    let breastKey = 'Breast';

    if (existingPatients && existingPatients.length > 0) {
      const firstPatientKeys = Object.keys(existingPatients[0].fields);

      const foundNameKey = firstPatientKeys.find(
        (k) => k.toLowerCase() === 'name' || k.toLowerCase().includes('patient name'),
      );
      if (foundNameKey) nameKey = foundNameKey;

      const foundDateKey = firstPatientKeys.find((k) => k.toLowerCase() === 'date' || k.toLowerCase().includes('referral date'));
      if (foundDateKey) dateKey = foundDateKey;

      const foundBreastKey = firstPatientKeys.find((k) => k.toLowerCase() === 'breast');
      if (foundBreastKey) breastKey = foundBreastKey;

      firstPatientKeys.forEach((k) => {
        if (k === nameKey) {
          finalFields[k] = name.trim();
        } else if (k === dateKey) {
          finalFields[k] = combinedDates;
        } else if (k === breastKey) {
          finalFields[k] = breast.trim();
        } else {
          finalFields[k] = (fieldValues[k] || '').trim();
        }
      });
    } else {
      finalFields['Referral Date'] = combinedDates;
      finalFields['Admit Number'] = admitNumber.trim();
      finalFields['name'] = name.trim();
      finalFields['Breast'] = breast.trim();
      Object.entries(fieldValues).forEach(([k, v]) => {
        finalFields[k] = (v || '').trim();
      });
    }

    finalFields['Admit Number'] = admitNumber.trim();

    customFields.forEach((cf) => {
      if (cf.label.trim()) {
        finalFields[cf.label.trim()] = cf.value.trim();
      }
    });

    onSave({
      patientName: name.trim(),
      date: primaryDate,
      breastValue: breast.trim(),
      fields: finalFields,
    });
  };

  const handleSave = () => {
    if (!name.trim()) {
      Alert.alert('Validation Error', 'Patient Name is required.');
      return;
    }

    const newDateValues = referralDates.map((d) => d.value.trim()).filter(Boolean);
    checkDuplicateDate(name.trim(), newDateValues, existingPatients, proceedWithSave);
  };

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
              <Text style={styles.sectionHeaderTitle}>Referral Dates</Text>
              <Pressable style={styles.addDateBtn} onPress={addReferralDate}>
                <Text style={styles.addDateBtnText}>+ Add Date</Text>
              </Pressable>
            </View>

            {referralDates.map((d, idx) => (
              <View style={styles.referralDateRow} key={d.id}>
                <Text style={styles.referralDateLabel}>Referral Date {idx + 1}</Text>
                <Pressable
                  style={[styles.formInput, { flex: 1, justifyContent: 'center' }]}
                  onPress={() => openDatePicker(d.id, d.value)}
                >
                  <Text style={{ color: '#1f2937', fontSize: 15 }}>{d.value}</Text>
                </Pressable>
                {referralDates.length > 1 && (
                  <Pressable
                    style={styles.customFieldRemoveBtn}
                    onPress={() => removeReferralDate(d.id)}
                  >
                    <Text style={styles.customFieldRemoveText}>X</Text>
                  </Pressable>
                )}
              </View>
            ))}

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
              <Text style={styles.formLabel}>Breast</Text>
              <TextInput
                style={styles.formInput}
                placeholder="e.g. UR, Warning"
                placeholderTextColor="#64748b"
                value={breast}
                onChangeText={setBreast}
              />
              <View style={styles.quickStatusRow}>
                <Pressable
                  style={[
                    styles.quickStatusBtn,
                    breast.toUpperCase() === 'UR' && styles.quickStatusBtnActiveUr,
                  ]}
                  onPress={() => setBreast('UR')}
                >
                  <Text style={[
                    styles.quickStatusBtnText,
                    breast.toUpperCase() === 'UR' && { color: '#065f46' }
                  ]}>Set UR</Text>
                </Pressable>
                <Pressable
                  style={[
                    styles.quickStatusBtn,
                    breast.toUpperCase() !== 'UR' && breast !== '' && styles.quickStatusBtnActiveNotUr,
                  ]}
                  onPress={() => setBreast('Warning')}
                >
                  <Text style={[
                    styles.quickStatusBtnText,
                    breast.toUpperCase() !== 'UR' && breast !== '' && { color: '#92400e' }
                  ]}>Warning</Text>
                </Pressable>
              </View>
            </View>

            <View style={styles.sectionHeaderRow}>
              <Text style={styles.sectionHeaderTitle}>Additional Details</Text>
            </View>

            {fieldKeys.map((key) => (
              <View style={styles.formGroup} key={key}>
                <Text style={styles.formLabel}>{key.replace(/\b\w/g, (c) => c.toUpperCase())}</Text>
                <TextInput
                  style={styles.formInput}
                  placeholder={`Enter ${key.toLowerCase()}`}
                  placeholderTextColor="#64748b"
                  value={fieldValues[key]}
                  onChangeText={(text) =>
                    setFieldValues((prev) => ({ ...prev, [key]: text }))
                  }
                />
              </View>
            ))}

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
  onClearData,
  hasData,
}: {
  visible: boolean;
  onClose: () => void;
  onAbout: () => void;
  onStaffs: () => void;
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
  const staffList = [
    { name: 'Sample', role: 'Breast Surgeon', dept: 'Breast Clinic' },
    { name: 'Sample', role: 'Radiologist', dept: 'Breast Imaging' },
    { name: 'Sample', role: 'Mammography Technologist', dept: 'Breast Imaging' },
    { name: 'Sample', role: 'Breast Care Nurse', dept: 'Breast Clinic' },
  ];

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
          <Text style={styles.staffsSubtitle}>Medical Team</Text>
          {staffList.map((staff, idx) => (
            <View key={idx} style={styles.staffCard}>
              <View style={styles.staffAvatar}>
                <Text style={styles.staffAvatarText}>{staff.name.charAt(0)}</Text>
              </View>
              <View style={styles.staffInfo}>
                <Text style={styles.staffName}>{staff.name}</Text>
                <Text style={styles.staffRole}>{staff.role}</Text>
                <Text style={styles.staffDept}>{staff.dept}</Text>
              </View>
            </View>
          ))}
        </ScrollView>
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
    marginBottom: 8,
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
    paddingHorizontal: 12,
    paddingVertical: 10,
    backgroundColor: '#ffe4e6',
    gap: 8,
  },
  listHeaderCell: {
    color: '#9f1239',
    fontSize: 12,
    fontWeight: '700',
    letterSpacing: 0.5,
  },
  listHeaderBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
  },
  listHeaderText: {
    color: '#9f1239',
    fontSize: 12,
    fontWeight: '700',
    letterSpacing: 0.5,
  },
  sortBadge: {
    backgroundColor: '#db4278',
    borderRadius: 4,
    paddingHorizontal: 5,
    paddingVertical: 1,
  },
  sortBadgeText: {
    color: '#ffffff',
    fontSize: 9,
    fontWeight: '700',
  },
  listRow: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: 12,
    paddingVertical: 12,
    borderBottomWidth: 1,
    borderBottomColor: '#f3e8ee',
    gap: 8,
    borderLeftWidth: 3,
    borderLeftColor: 'transparent',
  },
  listRowSelected: {
    backgroundColor: 'rgba(219, 66, 120, 0.08)',
    borderLeftColor: '#db4278',
  },
  listCell: {
    color: '#374151',
    fontSize: 14,
  },
  listDateCol: {
    flex: 1.1,
  },
  listNameCol: {
    flex: 1.4,
  },
  listNameText: {
    color: '#db4278',
    fontWeight: '600',
    textDecorationLine: 'underline',
  },
  listRecordCount: {
    color: '#9ca3af',
    fontSize: 11,
    marginTop: 2,
  },
  listStatusCol: {
    flex: 0.5,
    alignItems: 'flex-end',
  },
  statusBadgeWrap: {
    justifyContent: 'center',
  },
  statusBadge: {
    borderRadius: 999,
    paddingHorizontal: 10,
    paddingVertical: 4,
    minWidth: 36,
    alignItems: 'center',
  },
  statusBadgeOk: {
    backgroundColor: '#d1fae5',
  },
  statusBadgeWarn: {
    backgroundColor: '#fef3c7',
  },
  statusBadgeText: {
    fontSize: 14,
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
    marginBottom: 4,
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
});