import { useCallback, useState, useEffect, useMemo } from 'react';
import {
  ActivityIndicator,
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
import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';
import * as XLSX from 'xlsx';

import { ExcelParseResult, PatientRecord, parseExcelBase64 } from './src/utils/excel';

const teamSilayanLogo = require('./assets/team-silayan.png');

type LoadedFile = {
  name: string;
  result: ExcelParseResult;
};

export default function App() {
  const [loadedFile, setLoadedFile] = useState<LoadedFile | null>(null);
  const [selectedPatient, setSelectedPatient] = useState<PatientRecord | null>(null);
  const [loading, setLoading] = useState(false);
  const [aboutVisible, setAboutVisible] = useState(false);
  const [manualInputVisible, setManualInputVisible] = useState(false);
  const [editAdmitDatesPatient, setEditAdmitDatesPatient] = useState<PatientRecord | null>(null);

  const handleSaveManualPatient = useCallback(
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
          name: 'Manual Entries',
          result: {
            sheetName: 'Manual Sheet',
            patients: [newPatient],
          },
        });
      }

      setSelectedPatient(newPatient);
      setManualInputVisible(false);
    },
    [loadedFile],
  );

  const handleSaveAdmitDates = useCallback(
    (patient: PatientRecord, newDates: string[]) => {
      if (!loadedFile) return;
      const combinedDates = newDates.filter(Boolean).join('; ');
      const primaryDate = newDates[0] || '(no date)';

      // Find the key that represents the date in this patient's fields
      const fieldKeys = Object.keys(patient.fields);
      const dateKey = fieldKeys.find(
        (k) => k.toLowerCase() === 'date' || k.toLowerCase().includes('admit date'),
      ) || 'date';

      const updatedFields = { ...patient.fields, [dateKey]: combinedDates };
      const updatedPatient: PatientRecord = {
        ...patient,
        date: primaryDate,
        fields: updatedFields,
      };

      const updatedPatients = loadedFile.result.patients.map((p) =>
        p.rowIndex === patient.rowIndex ? updatedPatient : p,
      );
      setLoadedFile({
        ...loadedFile,
        result: { ...loadedFile.result, patients: updatedPatients },
      });
      // Keep the detail panel in sync
      setSelectedPatient(updatedPatient);
      setEditAdmitDatesPatient(null);
    },
    [loadedFile],
  );

  const exportExcel = useCallback(async () => {
    if (!loadedFile?.result.patients.length) {
      Alert.alert('Nothing to export', 'Load or create some patient records first.');
      return;
    }
    try {
      const patients = loadedFile.result.patients;
      // Build header row from first patient fields
      const headers = Object.keys(patients[0].fields);
      const rows = patients.map((p) => headers.map((h) => p.fields[h] ?? ''));
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

  const importExcel = useCallback(async () => {
    try {
      setLoading(true);
      setSelectedPatient(null);

      const result = await DocumentPicker.getDocumentAsync({
        type: [
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'application/vnd.ms-excel',
        ],
        copyToCacheDirectory: true,
      });

      if (result.canceled || !result.assets?.[0]) {
        return;
      }

      const asset = result.assets[0];
      const base64 = await FileSystem.readAsStringAsync(asset.uri, {
        encoding: FileSystem.EncodingType.Base64,
      });

      const checkResult = parseExcelBase64(base64);

      // Inject admit number for any patient that doesn't have one
      const admitNumberKey = Object.keys(checkResult.patients[0]?.fields ?? {}).find(
        (k) => k.toLowerCase().includes('admit number')
      );

      if (!admitNumberKey) {
        checkResult.patients.forEach((patient) => {
          patient.fields['Admit Number'] = generateAdmitNumber();
        });
      }

      setLoadedFile({
        name: asset.name,
        result: checkResult,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to read the Excel file.';
      Alert.alert('Import failed', message);
    } finally {
      setLoading(false);
    }
  }, []);

  const { result } = loadedFile ?? {};

  return (
    <SafeAreaView style={styles.safeArea}>
      <StatusBar style="light" />
      <View style={styles.container}>
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
            <Pressable style={styles.aboutButton} onPress={() => setAboutVisible(true)}>
              <Text style={styles.aboutButtonText}>About</Text>
            </Pressable>
          </View>
          <Text style={styles.subtitle}>
            Import a clinical Excel sheet. Each row is listed with date, name, and Breast status.
            Tap a name to view full details.
          </Text>
        </View>

        <View style={styles.buttonRow}>
          <Pressable
            style={[styles.button, styles.primaryButton, styles.rowButton, loading && styles.buttonDisabled]}
            onPress={importExcel}
            disabled={loading}
          >
            <LinearGradient
              colors={loading ? ['#f3a5c2', '#db4278'] : ['#db4278', '#b51f55']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 0 }}
              style={StyleSheet.absoluteFillObject}
            />
            {loading ? (
              <ActivityIndicator color="#ffffff" />
            ) : (
              <Text style={styles.primaryButtonText}>Import Excel</Text>
            )}
          </Pressable>

          <Pressable
            style={[styles.button, styles.secondaryButton, styles.rowButton]}
            onPress={() => setManualInputVisible(true)}
          >
            <Text style={styles.secondaryButtonText}>Manual Input</Text>
          </Pressable>

          <Pressable
            style={[styles.button, styles.exportButton, styles.rowButton]}
            onPress={exportExcel}
          >
            <Text style={styles.exportButtonText}>Export</Text>
          </Pressable>
        </View>

        {loadedFile && result ? (
          <ScrollView style={styles.resultScroll} contentContainerStyle={styles.resultSection}>
            <View style={styles.fileCard}>
              <Text style={styles.fileName}>{loadedFile.name}</Text>
              <Text style={styles.fileMeta}>
                {result.patients.length} patient{result.patients.length === 1 ? '' : 's'} ·{' '}
                {result.sheetName}
              </Text>
            </View>

            <View style={styles.listCard}>
              <View style={styles.listHeader}>
                <Text style={[styles.listHeaderCell, styles.listDateCol]}>Admit Date</Text>
                <Text style={[styles.listHeaderCell, styles.listNameCol]}>Name</Text>
                <Text style={[styles.listHeaderCell, styles.listStatusCol]}>Status</Text>
              </View>

              {result.patients.map((patient) => (
                <PatientListRow
                  key={patient.rowIndex}
                  patient={patient}
                  selected={selectedPatient?.rowIndex === patient.rowIndex}
                  onPress={() =>
                    setSelectedPatient((current) =>
                      current?.rowIndex === patient.rowIndex ? null : patient,
                    )
                  }
                />
              ))}
            </View>

            {selectedPatient ? (
              <PatientDetail
                patient={selectedPatient}
                onEditAdmitDates={(p) => setEditAdmitDatesPatient(p)}
              />
            ) : null}
          </ScrollView>
        ) : (
          <View style={styles.placeholderCard}>
            <Text style={styles.placeholderTitle}>How it works</Text>
            <Text style={styles.placeholderText}>
              1. Import your clinical Excel file or tap "Manual Input" to create custom records
            </Text>
            <Text style={styles.placeholderText}>2. All data rows appear in the list</Text>
            <Text style={styles.placeholderText}>3. Each row shows date, name, and status</Text>
            <Text style={styles.placeholderText}>4. Tap a name to see full row details</Text>
          </View>
        )}
      </View>

      <AboutModal visible={aboutVisible} onClose={() => setAboutVisible(false)} />
      <ManualInputModal
        visible={manualInputVisible}
        onClose={() => setManualInputVisible(false)}
        onSave={handleSaveManualPatient}
        existingPatients={loadedFile?.result.patients}
      />
      {editAdmitDatesPatient && (
        <EditAdmitDatesModal
          patient={editAdmitDatesPatient}
          onClose={() => setEditAdmitDatesPatient(null)}
          onSave={handleSaveAdmitDates}
        />
      )}
    </SafeAreaView>
  );
}

function PatientListRow({
  patient,
  selected,
  onPress,
}: {
  patient: PatientRecord;
  selected: boolean;
  onPress: () => void;
}) {
  return (
    <View style={[styles.listRow, selected && styles.listRowSelected]}>
      <Text style={[styles.listCell, styles.listDateCol]} numberOfLines={1}>
        {patient.date}
      </Text>
      <Pressable style={styles.listNameCol} onPress={onPress}>
        <Text style={[styles.listCell, styles.listNameText]} numberOfLines={1}>
          {patient.patientName}
        </Text>
      </Pressable>
      <View style={[styles.listStatusCol, styles.statusBadgeWrap]}>
        <View style={[styles.statusBadge, patient.isUr ? styles.statusBadgeOk : styles.statusBadgeWarn]}>
          <Text style={[styles.statusBadgeText, { color: patient.isUr ? '#065f46' : '#92400e' }]}>
            {patient.isUr ? '✓ UR' : '⚠ Warning'}
          </Text>
        </View>
      </View>
    </View>
  );
}

function PatientDetail({
  patient,
  onEditAdmitDates,
}: {
  patient: PatientRecord;
  onEditAdmitDates: (p: PatientRecord) => void;
}) {
  return (
    <View style={styles.detailSection}>
      <Text style={styles.detailSectionTitle}>{patient.patientName}</Text>
      <Text style={styles.detailSectionMeta}>Row {patient.rowIndex}</Text>

      <View
        style={[
          styles.statusCard,
          patient.isUr ? styles.statusCardOk : styles.statusCardWarn,
        ]}
      >
        <LinearGradient
          colors={patient.isUr ? ['#eafaf1', '#d1f2e1'] : ['#fdf6e2', '#f9e8be']}
          start={{ x: 0, y: 0 }}
          end={{ x: 1, y: 1 }}
          style={StyleSheet.absoluteFillObject}
        />
        <Text style={[styles.statusIcon, { color: patient.isUr ? '#10b981' : '#f59e0b' }]}>
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

      {/* Admit Dates panel */}
      <View style={styles.admitDatesPanelCard}>
        <View style={styles.admitDatesPanelHeader}>
          <Text style={styles.admitDatesPanelTitle}>📅 Admit Dates</Text>
          <Pressable
            style={styles.editAdmitDatesBtn}
            onPress={() => onEditAdmitDates(patient)}
          >
            <Text style={styles.editAdmitDatesBtnText}>Edit</Text>
          </Pressable>
        </View>
        {patient.date && patient.date !== '(no date)' ? (
          patient.date.split(';').map((d, i) => (
            <Text key={i} style={styles.admitDatesPanelEntry}>
              {i + 1}. {d.trim()}
            </Text>
          ))
        ) : (
          <Text style={styles.admitDatesPanelEmpty}>No admit dates recorded</Text>
        )}
      </View>

      <View style={styles.detailCard}>
        {Object.entries(patient.fields).map(([label, value]) => (
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
      <Text style={styles.detailLabel}>{label}</Text>
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

function EditAdmitDatesModal({
  patient,
  onClose,
  onSave,
}: {
  patient: PatientRecord;
  onClose: () => void;
  onSave: (patient: PatientRecord, dates: string[]) => void;
}) {
  // Parse the existing date string (semicolon-separated) into an array
  const parsedInitial = useMemo(() => {
    const raw = patient.date || '';
    if (!raw || raw === '(no date)') return [{ id: 'init', value: todayString() }];
    return raw.split(';').map((d, i) => ({ id: `existing-${i}`, value: d.trim() }));
  }, [patient.date]);

  const [dates, setDates] = useState(parsedInitial);

  // Re-sync when patient changes
  useEffect(() => { setDates(parsedInitial); }, [parsedInitial]);

  const addDate = () =>
    setDates((prev) => [...prev, { id: Math.random().toString(36).slice(2), value: todayString() }]);

  const updateDate = (id: string, value: string) =>
    setDates((prev) => prev.map((d) => (d.id === id ? { ...d, value } : d)));

  const removeDate = (id: string) =>
    setDates((prev) => prev.filter((d) => d.id !== id));

  const handleSave = () => {
    const values = dates.map((d) => d.value.trim()).filter(Boolean);
    onSave(patient, values.length ? values : ['(no date)']);
  };

  return (
    <Modal visible animationType="slide" transparent={false} onRequestClose={onClose}>
      <SafeAreaView style={styles.modalContainer}>
        <KeyboardAvoidingView behavior={Platform.OS === 'ios' ? 'padding' : 'height'} style={{ flex: 1 }}>
          <View style={styles.modalHeaderRow}>
            <Text style={styles.modalHeaderTitle}>Edit Admit Dates</Text>
            <Pressable style={styles.modalCancelButton} onPress={onClose}>
              <Text style={styles.modalCancelButtonText}>Cancel</Text>
            </Pressable>
          </View>

          <ScrollView style={styles.modalScrollView} contentContainerStyle={styles.modalScrollContent}>
            <Text style={styles.editAdmitPatientName}>{patient.patientName}</Text>
            <Text style={styles.editAdmitPatientMeta}>Row {patient.rowIndex}</Text>

            <View style={[styles.sectionHeaderRow, { marginTop: 20 }]}>
              <Text style={styles.sectionHeaderTitle}>Admit Dates</Text>
              <Pressable style={styles.addDateBtn} onPress={addDate}>
                <Text style={styles.addDateBtnText}>+ Add Date</Text>
              </Pressable>
            </View>

            {dates.map((d, idx) => (
              <View style={styles.admitDateRow} key={d.id}>
                <Text style={styles.admitDateLabel}>Date {idx + 1}</Text>
                <TextInput
                  style={[styles.formInput, { flex: 1 }]}
                  placeholder="MM/DD/YYYY"
                  placeholderTextColor="#64748b"
                  value={d.value}
                  onChangeText={(text) => updateDate(d.id, text)}
                />
                {dates.length > 1 && (
                  <Pressable style={styles.customFieldRemoveBtn} onPress={() => removeDate(d.id)}>
                    <Text style={styles.customFieldRemoveText}>✕</Text>
                  </Pressable>
                )}
              </View>
            ))}

            <Pressable style={[styles.saveBtn, { marginTop: 32 }]} onPress={handleSave}>
              <LinearGradient
                colors={['#db4278', '#b51f55']}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 0 }}
                style={StyleSheet.absoluteFillObject}
              />
              <Text style={styles.saveBtnText}>Save Admit Dates</Text>
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

function ManualInputModal({
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
  const [admitDates, setAdmitDates] = useState<Array<{ id: string; value: string }>>([]);
  const [admitNumber, setAdmitNumber] = useState('');
  const [breast, setBreast] = useState('UR');
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const [customFields, setCustomFields] = useState<Array<{ id: string; label: string; value: string }>>([]);

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
      setAdmitDates([{ id: 'init', value: todayString() }]);
      setAdmitNumber(generateAdmitNumber());
      setBreast('UR');
      const initialValues: Record<string, string> = {};
      fieldKeys.forEach((key) => { initialValues[key] = ''; });
      setFieldValues(initialValues);
      setCustomFields([]);
    }
  }, [visible, fieldKeys]);

  const addAdmitDate = () => {
    setAdmitDates((prev) => [
      ...prev,
      { id: Math.random().toString(36).substring(2, 9), value: todayString() },
    ]);
  };

  const updateAdmitDate = (id: string, value: string) => {
    setAdmitDates((prev) => prev.map((d) => (d.id === id ? { ...d, value } : d)));
  };

  const removeAdmitDate = (id: string) => {
    setAdmitDates((prev) => prev.filter((d) => d.id !== id));
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

  const handleSave = () => {
    if (!name.trim()) {
      Alert.alert('Validation Error', 'Patient Name is required.');
      return;
    }

    // Combine all admit dates into one string (semicolon-separated)
    const combinedDates = admitDates
      .map((d) => d.value.trim())
      .filter(Boolean)
      .join('; ');
    const primaryDate = admitDates[0]?.value.trim() || '(no date)';

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

      const foundDateKey = firstPatientKeys.find((k) => k.toLowerCase() === 'date' || k.toLowerCase().includes('admit date'));
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
      finalFields['Admit Date'] = combinedDates;
      finalFields['Admit Number'] = admitNumber.trim();
      finalFields['name'] = name.trim();
      finalFields['Breast'] = breast.trim();
      Object.entries(fieldValues).forEach(([k, v]) => {
        finalFields[k] = (v || '').trim();
      });
    }

    // Always store admit number
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
                  <Text style={styles.regenBtnText}>↻ New</Text>
                </Pressable>
              </View>
            </View>

            <View style={styles.sectionHeaderRow}>
              <Text style={styles.sectionHeaderTitle}>Admit Dates</Text>
              <Pressable style={styles.addDateBtn} onPress={addAdmitDate}>
                <Text style={styles.addDateBtnText}>+ Add Date</Text>
              </Pressable>
            </View>

            {admitDates.map((d, idx) => (
              <View style={styles.admitDateRow} key={d.id}>
                <Text style={styles.admitDateLabel}>Admit Date {idx + 1}</Text>
                <TextInput
                  style={[styles.formInput, { flex: 1 }]}
                  placeholder="MM/DD/YYYY"
                  placeholderTextColor="#64748b"
                  value={d.value}
                  onChangeText={(text) => updateAdmitDate(d.id, text)}
                />
                {admitDates.length > 1 && (
                  <Pressable
                    style={styles.customFieldRemoveBtn}
                    onPress={() => removeAdmitDate(d.id)}
                  >
                    <Text style={styles.customFieldRemoveText}>✕</Text>
                  </Pressable>
                )}
              </View>
            ))}

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
                  ]}>✓ Set UR</Text>
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
                  ]}>⚠ Warning</Text>
                </Pressable>
              </View>
            </View>

            <View style={styles.sectionHeaderRow}>
              <Text style={styles.sectionHeaderTitle}>Additional Details</Text>
            </View>

            {fieldKeys.map((key) => (
              <View style={styles.formGroup} key={key}>
                <Text style={styles.formLabel}>{key}</Text>
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
                  <Text style={styles.customFieldRemoveText}>✕</Text>
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
    backgroundColor: 'rgba(0, 0, 0, 0.25)',
    borderRadius: 8,
    padding: 8,
    marginHorizontal: -20,   // add this (match your header padding value)
    marginBottom: -20,       // add this
    paddingHorizontal: 20,   // add this
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
    fontSize: 13,
    fontWeight: '700',
    textTransform: 'uppercase',
    letterSpacing: 0.5,
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
  listStatusCol: {
    flex: 1,
    alignItems: 'flex-end',
  },
  statusBadgeWrap: {
    justifyContent: 'center',
  },
  statusBadge: {
    borderRadius: 999,
    paddingHorizontal: 10,
    paddingVertical: 4,
  },
  statusBadgeOk: {
    backgroundColor: '#d1fae5',
  },
  statusBadgeWarn: {
    backgroundColor: '#fef3c7',
  },
  statusBadgeText: {
    fontSize: 12,
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
  admitDateRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
  },
  admitDateLabel: {
    color: '#6b7280',
    fontSize: 13,
    fontWeight: '600',
    width: 80,
  },
  admitDatesPanelCard: {
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
  admitDatesPanelHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: 4,
  },
  admitDatesPanelTitle: {
    color: '#1f2937',
    fontSize: 15,
    fontWeight: '700',
  },
  editAdmitDatesBtn: {
    paddingHorizontal: 14,
    paddingVertical: 6,
    borderRadius: 8,
    backgroundColor: '#fff1f5',
    borderWidth: 1,
    borderColor: '#db4278',
  },
  editAdmitDatesBtnText: {
    color: '#db4278',
    fontSize: 13,
    fontWeight: '700',
  },
  admitDatesPanelEntry: {
    color: '#374151',
    fontSize: 14,
    lineHeight: 22,
  },
  admitDatesPanelEmpty: {
    color: '#9ca3af',
    fontSize: 14,
    fontStyle: 'italic',
  },
  editAdmitPatientName: {
    color: '#1f2937',
    fontSize: 20,
    fontWeight: '700',
  },
  editAdmitPatientMeta: {
    color: '#6b7280',
    fontSize: 14,
    marginTop: 2,
  },
});
