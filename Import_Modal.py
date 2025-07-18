'use client';

import { useState, useRef, useEffect } from 'react';
import { XMarkIcon, DocumentArrowUpIcon, DocumentTextIcon, TableCellsIcon, ClipboardDocumentIcon, ArrowDownTrayIcon, CheckCircleIcon, ExclamationTriangleIcon, TrashIcon } from '@heroicons/react/24/outline';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

export default function ImportModal({ isOpen, onClose, onImport, existingAutomations = [] }) {
  const [step, setStep] = useState(1); // 1: Select File, 2: Preview, 3: Validation Results
  const [selectedFile, setSelectedFile] = useState(null);
  const [fileType, setFileType] = useState(null);
  const [parsedData, setParsedData] = useState([]);
  const [previewData, setPreviewData] = useState([]);
  const [importPreview, setImportPreview] = useState({ 
    newRecords: [], 
    updateRecords: [], 
    deleteRecords: [], 
    errors: [] 
  });
  const [isProcessing, setIsProcessing] = useState(false);
  const [importProgress, setImportProgress] = useState({ current: 0, total: 0 });
  const [showTemplate, setShowTemplate] = useState(false);
  const fileInputRef = useRef(null);

  // Template data structure based on the Automation model
  const templateHeaders = [
    'air_id', 'name', 'type', 'brief_description', 'coe_fed', 'complexity',
    'tool_version', 'process_details', 'object_details', 'queue',
    'shared_folders', 'shared_mailboxes', 'qa_handshake',
    'preprod_deploy_date', 'prod_deploy_date', 'warranty_end_date',
    'comments', 'documentation', 'modified', 'path'
  ];

  const sampleData = [
    {
      air_id: 'AIR-2024-001',
      name: 'Invoice Processing Automation',
      type: 'RPA',
      brief_description: 'Automated invoice processing and validation',
      coe_fed: 'Finance',
      complexity: 'Medium',
      tool_version: 'UiPath 2023.4',
      process_details: 'Processes invoices from shared mailbox',
      object_details: 'Invoice data extraction and validation',
      queue: 'InvoiceQueue',
      shared_folders: '\\\\server\\invoices',
      shared_mailboxes: 'invoices@company.com',
      qa_handshake: 'QA-001',
      preprod_deploy_date: '2024-01-15',
      prod_deploy_date: '2024-02-01',
      warranty_end_date: '2024-08-01',
      comments: 'Initial version deployed',
      documentation: 'https://docs.company.com/automation-001',
      modified: '2024-01-20',
      path: '\\\\automation\\invoices'
    }
  ];

  // Reset modal state when opening/closing
  useEffect(() => {
    if (isOpen) {
      setStep(1);
      setSelectedFile(null);
      setFileType(null);
      setParsedData([]);
      setPreviewData([]);
      setImportPreview({ newRecords: [], updateRecords: [], deleteRecords: [], errors: [] });
      setIsProcessing(false);
      setImportProgress({ current: 0, total: 0 });
      setShowTemplate(false);
    }
  }, [isOpen]);

  const downloadTemplate = (format) => {
    const ws = XLSX.utils.json_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Automation Template');

    if (format === 'excel') {
      XLSX.writeFile(wb, 'automation_import_template.xlsx');
    } else if (format === 'csv') {
      const csv = Papa.unparse(sampleData);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'automation_import_template.csv';
      link.click();
    } else if (format === 'json') {
      const jsonData = JSON.stringify(sampleData, null, 2);
      const blob = new Blob([jsonData], { type: 'application/json;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'automation_import_template.json';
      link.click();
    }
  };

  const handleFileSelect = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    setSelectedFile(file);
    setIsProcessing(true);

    try {
      const extension = file.name.split('.').pop().toLowerCase();
      let data = [];

      if (extension === 'csv') {
        setFileType('CSV');
        const text = await file.text();
        const result = Papa.parse(text, {
          header: true,
          skipEmptyLines: true,
          transform: (value) => value.trim()
        });
        data = result.data;
      } else if (['xlsx', 'xls'].includes(extension)) {
        setFileType('Excel');
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        data = XLSX.utils.sheet_to_json(worksheet);
      } else if (extension === 'json') {
        setFileType('JSON');
        const text = await file.text();
        data = JSON.parse(text);
      } else {
        throw new Error('Unsupported file format. Please use CSV, Excel, or JSON files.');
      }

      setParsedData(data);
      setPreviewData(data.slice(0, 10)); // Show first 10 rows for preview
      analyzeImportData(data);
      setStep(2);
    } catch (error) {
      alert(`Error parsing file: ${error.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // Helper function to normalize values for comparison
  const normalizeValue = (value) => {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'string') return value.trim();
    return value;
  };

  // Helper function to check if two records have different values
  const hasChanges = (newRecord, existingRecord) => {
    for (const field of templateHeaders) {
      const newValue = normalizeValue(newRecord[field]);
      const existingValue = normalizeValue(existingRecord[field]);
      
      // Special handling for dates
      if (['preprod_deploy_date', 'prod_deploy_date', 'warranty_end_date', 'modified'].includes(field)) {
        if (newValue && existingValue) {
          const newDate = new Date(newValue).toISOString().split('T')[0];
          const existingDate = new Date(existingValue).toISOString().split('T')[0];
          if (newDate !== existingDate) return true;
        } else if (newValue !== existingValue) {
          return true;
        }
      } else if (newValue !== existingValue) {
        return true;
      }
    }
    return false;
  };

  const analyzeImportData = (data) => {
    const newRecords = [];
    const updateRecords = [];
    const deleteRecords = [];
    const errors = [];
    const existingAirIds = new Set(existingAutomations.map(a => a.air_id));
    const importAirIds = new Set();

    // Process import data
    data.forEach((row, index) => {
      const validation = validateRow(row, index + 1);
      if (validation.errors.length > 0) {
        errors.push(...validation.errors);
        return;
      }

      const cleanedRow = validation.data;
      importAirIds.add(cleanedRow.air_id);

      if (existingAirIds.has(cleanedRow.air_id)) {
        const existingRecord = existingAutomations.find(a => a.air_id === cleanedRow.air_id);
        
        // Only add to update list if there are actual changes
        if (hasChanges(cleanedRow, existingRecord)) {
          updateRecords.push({
            ...cleanedRow,
            existingData: existingRecord,
            changes: getChangedFields(cleanedRow, existingRecord)
          });
        }
      } else {
        newRecords.push(cleanedRow);
      }
    });

    // Find records to delete (existing records not in import file)
    existingAutomations.forEach(existingRecord => {
      if (!importAirIds.has(existingRecord.air_id)) {
        deleteRecords.push(existingRecord);
      }
    });

    setImportPreview({ newRecords, updateRecords, deleteRecords, errors });
  };

  const validateRow = (row, rowNumber) => {
    const errors = [];
    const data = {};

    // Required fields validation
    if (!row.air_id || !row.air_id.trim()) {
      errors.push(`Row ${rowNumber}: AIR ID is required`);
    } else {
      data.air_id = row.air_id.trim();
    }

    if (!row.name || !row.name.trim()) {
      errors.push(`Row ${rowNumber}: Name is required`);
    } else {
      data.name = row.name.trim();
    }

    if (!row.type || !row.type.trim()) {
      errors.push(`Row ${rowNumber}: Type is required`);
    } else {
      data.type = row.type.trim();
    }

    // Map all template headers with proper null handling
    templateHeaders.forEach(field => {
      if (!['air_id', 'name', 'type'].includes(field)) {
        data[field] = row[field] ? row[field].toString().trim() : null;
        if (data[field] === '') data[field] = null;
      }
    });

    // Date validation and formatting
    ['preprod_deploy_date', 'prod_deploy_date', 'warranty_end_date', 'modified'].forEach(dateField => {
      if (data[dateField]) {
        try {
          const date = new Date(data[dateField]);
          if (isNaN(date.getTime())) {
            errors.push(`Row ${rowNumber}: Invalid date format for ${dateField}`);
            data[dateField] = null;
          } else {
            data[dateField] = date.toISOString();
          }
        } catch (e) {
          errors.push(`Row ${rowNumber}: Invalid date format for ${dateField}`);
          data[dateField] = null;
        }
      }
    });

    return { data, errors };
  };

  const handleImport = async () => {
    setIsProcessing(true);
    setStep(3);

    const totalOperations = importPreview.newRecords.length + 
                           importPreview.updateRecords.length + 
                           importPreview.deleteRecords.length;
    let currentOperation = 0;
    let successCount = 0;
    let errorCount = 0;
    const importErrors = [];

    setImportProgress({ current: 0, total: totalOperations });

    // Process new records
    for (const record of importPreview.newRecords) {
      currentOperation++;
      setImportProgress({ current: currentOperation, total: totalOperations });

      try {
        const response = await fetch('/api/automations', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(record),
        });

        if (response.ok) {
          successCount++;
        } else {
          errorCount++;
          const errorText = await response.text();
          importErrors.push(`Create ${record.air_id}: ${errorText}`);
        }
      } catch (error) {
        errorCount++;
        importErrors.push(`Create ${record.air_id}: ${error.message}`);
      }
    }

    // Process updates (only records with changes)
    for (const record of importPreview.updateRecords) {
      currentOperation++;
      setImportProgress({ current: currentOperation, total: totalOperations });

      try {
        const response = await fetch(`/api/automations/${record.air_id}`, {
          method: 'PUT',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(record),
        });

        if (response.ok) {
          successCount++;
        } else {
          errorCount++;
          const errorText = await response.text();
          importErrors.push(`Update ${record.air_id}: ${errorText}`);
        }
      } catch (error) {
        errorCount++;
        importErrors.push(`Update ${record.air_id}: ${error.message}`);
      }
    }

    // Process deletions
    for (const record of importPreview.deleteRecords) {
      currentOperation++;
      setImportProgress({ current: currentOperation, total: totalOperations });

      try {
        const response = await fetch(`/api/automations/${record.air_id}`, {
          method: 'DELETE',
        });

        if (response.ok) {
          successCount++;
        } else {
          errorCount++;
          const errorText = await response.text();
          importErrors.push(`Delete ${record.air_id}: ${errorText}`);
        }
      } catch (error) {
        errorCount++;
        importErrors.push(`Delete ${record.air_id}: ${error.message}`);
      }
    }

    // Call the parent's onImport callback to refresh data
    if (onImport) {
      await onImport();
    }

    // Show final results
    const message = `Import completed!\nCreated: ${importPreview.newRecords.length}\nUpdated: ${importPreview.updateRecords.length}\nDeleted: ${importPreview.deleteRecords.length}\nSuccessful: ${successCount}\nErrors: ${errorCount}`;
    if (importErrors.length > 0) {
      console.error('Import errors:', importErrors);
    }

    setIsProcessing(false);

    // Auto-close after a delay if successful
    if (errorCount === 0) {
      setTimeout(() => onClose(), 3000);
    }
  };

  const getChangedFields = (newData, existingData) => {
    const changes = {};
    templateHeaders.forEach(field => {
      const newValue = normalizeValue(newData[field]);
      const existingValue = normalizeValue(existingData[field]);
      
      if (newValue !== existingValue) {
        changes[field] = { old: existingValue, new: newValue };
      }
    });
    return changes;
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg shadow-xl max-w-6xl w-full max-h-[90vh] overflow-hidden">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-200 flex items-center justify-between">
          <h2 className="text-xl font-semibold text-gray-900">
            Import Automations - Step {step} of 3
          </h2>
          <button
            onClick={onClose}
            className="text-gray-400 hover:text-gray-600"
          >
            <XMarkIcon className="h-6 w-6" />
          </button>
        </div>

        {/* Content */}
        <div className="px-6 py-4 max-h-[calc(90vh-120px)] overflow-y-auto">
          {step === 1 && (
            <div className="space-y-6">
              {/* Template Download Section */}
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                <h3 className="text-lg font-medium text-blue-900 mb-3">
                  Download Import Template
                </h3>
                <p className="text-blue-700 mb-4">
                  Download a template file with the correct column headers and sample data to ensure your import succeeds.
                </p>
                <div className="flex space-x-3">
                  <button
                    onClick={() => downloadTemplate('excel')}
                    className="flex items-center px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700"
                  >
                    <TableCellsIcon className="h-4 w-4 mr-2" />
                    Excel Template
                  </button>
                  <button
                    onClick={() => downloadTemplate('csv')}
                    className="flex items-center px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700"
                  >
                    <DocumentTextIcon className="h-4 w-4 mr-2" />
                    CSV Template
                  </button>
                  <button
                    onClick={() => downloadTemplate('json')}
                    className="flex items-center px-4 py-2 bg-purple-600 text-white rounded-md hover:bg-purple-700"
                  >
                    <ClipboardDocumentIcon className="h-4 w-4 mr-2" />
                    JSON Template
                  </button>
                </div>
              </div>

              {/* File Upload Section */}
              <div className="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center">
                <DocumentArrowUpIcon className="h-12 w-12 text-gray-400 mx-auto mb-4" />
                <h3 className="text-lg font-medium text-gray-900 mb-2">
                  Select Import File
                </h3>
                <p className="text-gray-600 mb-4">
                  Choose a CSV, Excel (.xlsx/.xls), or JSON file to import automation data
                </p>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".csv,.xlsx,.xls,.json"
                  onChange={handleFileSelect}
                  className="hidden"
                />
                <button
                  onClick={() => fileInputRef.current?.click()}
                  disabled={isProcessing}
                  className="px-6 py-3 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 disabled:opacity-50"
                >
                  {isProcessing ? 'Processing...' : 'Choose File'}
                </button>
              </div>

              {/* Supported Fields */}
              <div className="bg-gray-50 rounded-lg p-4">
                <h4 className="font-medium text-gray-900 mb-2">Supported Fields</h4>
                <div className="grid grid-cols-3 gap-2 text-sm text-gray-600">
                  {templateHeaders.map(header => (
                    <div key={header} className="px-2 py-1 bg-white rounded">
                      {header} {['air_id', 'name', 'type'].includes(header) && <span className="text-red-500">*</span>}
                    </div>
                  ))}
                </div>
                <p className="text-xs text-gray-500 mt-2">* Required fields</p>
              </div>
            </div>
          )}

          {step === 2 && (
            <div className="space-y-6">
              {/* File Info */}
              <div className="bg-gray-50 rounded-lg p-4">
                <h3 className="font-medium text-gray-900 mb-2">File Information</h3>
                <div className="grid grid-cols-3 gap-4 text-sm">
                  <div>
                    <span className="text-gray-600">File: </span>
                    <span className="font-medium">{selectedFile?.name}</span>
                  </div>
                  <div>
                    <span className="text-gray-600">Type: </span>
                    <span className="font-medium">{fileType}</span>
                  </div>
                  <div>
                    <span className="text-gray-600">Records: </span>
                    <span className="font-medium">{parsedData.length}</span>
                  </div>
                </div>
              </div>

              {/* Import Summary */}
              <div className="grid grid-cols-4 gap-4">
                <div className="bg-green-50 border border-green-200 rounded-lg p-4">
                  <div className="flex items-center">
                    <CheckCircleIcon className="h-8 w-8 text-green-600 mr-3" />
                    <div>
                      <div className="text-2xl font-bold text-green-900">
                        {importPreview.newRecords.length}
                      </div>
                      <div className="text-sm text-green-700">New Records</div>
                    </div>
                  </div>
                </div>
                <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                  <div className="flex items-center">
                    <DocumentTextIcon className="h-8 w-8 text-blue-600 mr-3" />
                    <div>
                      <div className="text-2xl font-bold text-blue-900">
                        {importPreview.updateRecords.length}
                      </div>
                      <div className="text-sm text-blue-700">Updates</div>
                    </div>
                  </div>
                </div>
                <div className="bg-red-50 border border-red-200 rounded-lg p-4">
                  <div className="flex items-center">
                    <TrashIcon className="h-8 w-8 text-red-600 mr-3" />
                    <div>
                      <div className="text-2xl font-bold text-red-900">
                        {importPreview.deleteRecords.length}
                      </div>
                      <div className="text-sm text-red-700">Deletions</div>
                    </div>
                  </div>
                </div>
                <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
                  <div className="flex items-center">
                    <ExclamationTriangleIcon className="h-8 w-8 text-yellow-600 mr-3" />
                    <div>
                      <div className="text-2xl font-bold text-yellow-900">
                        {importPreview.errors.length}
                      </div>
                      <div className="text-sm text-yellow-700">Errors</div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Error List */}
              {importPreview.errors.length > 0 && (
                <div className="bg-red-50 border border-red-200 rounded-lg p-4">
                  <h4 className="font-medium text-red-900 mb-2">Validation Errors</h4>
                  <div className="max-h-32 overflow-y-auto">
                    {importPreview.errors.map((error, index) => (
                      <div key={index} className="text-sm text-red-700 mb-1">
                        {error}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Records to Delete */}
              {importPreview.deleteRecords.length > 0 && (
                <div className="bg-red-50 border border-red-200 rounded-lg p-4">
                  <h4 className="font-medium text-red-900 mb-2">
                    Records to be Deleted (not in import file)
                  </h4>
                  <div className="max-h-32 overflow-y-auto">
                    {importPreview.deleteRecords.map((record, index) => (
                      <div key={index} className="text-sm text-red-700 mb-1">
                        {record.air_id} - {record.name}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Data Preview */}
              <div className="bg-white border border-gray-200 rounded-lg overflow-hidden">
                <div className="px-4 py-3 bg-gray-50 border-b border-gray-200">
                  <h4 className="font-medium text-gray-900">Data Preview (First 10 Records)</h4>
                </div>
                <div className="overflow-x-auto">
                  <table className="min-w-full divide-y divide-gray-200">
                    <thead className="bg-gray-50">
                      <tr>
                        {templateHeaders.slice(0, 6).map(header => (
                          <th key={header} className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                            {header}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-gray-200">
                      {previewData.map((row, index) => (
                        <tr key={index}>
                          {templateHeaders.slice(0, 6).map(header => (
                            <td key={header} className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                              {row[header] || '-'}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {step === 3 && (
            <div className="space-y-6">
              {/* Progress */}
              <div className="bg-gray-50 rounded-lg p-4">
                <h3 className="font-medium text-gray-900 mb-2">Import Progress</h3>
                <div className="w-full bg-gray-200 rounded-full h-2">
                  <div 
                    className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                    style={{ width: `${(importProgress.current / importProgress.total) * 100}%` }}
                  ></div>
                </div>
                <p className="text-sm text-gray-600 mt-2">
                  {importProgress.current} of {importProgress.total} operations completed
                </p>
              </div>

              {/* Final Results */}
              {!isProcessing && (
                <div className="bg-green-50 border border-green-200 rounded-lg p-4">
                  <h3 className="font-medium text-green-900 mb-2">Import Complete</h3>
                  <div className="grid grid-cols-4 gap-4 text-sm">
                    <div>
                      <span className="text-green-700">Created: </span>
                      <span className="font-medium">{importPreview.newRecords.length}</span>
                    </div>
                    <div>
                      <span className="text-green-700">Updated: </span>
                      <span className="font-medium">{importPreview.updateRecords.length}</span>
                    </div>
                    <div>
                      <span className="text-green-700">Deleted: </span>
                      <span className="font-medium">{importPreview.deleteRecords.length}</span>
                    </div>
                    <div>
                      <span className="text-green-700">Total: </span>
                      <span className="font-medium">{importProgress.total}</span>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-gray-200 flex items-center justify-between">
          <div className="flex space-x-3">
            {step > 1 && (
              <button
                onClick={() => setStep(step - 1)}
                disabled={isProcessing}
                className="px-4 py-2 text-gray-700 bg-gray-100 rounded-md hover:bg-gray-200 disabled:opacity-50"
              >
                Back
              </button>
            )}
          </div>
          <div className="flex space-x-3">
            <button
              onClick={onClose}
              disabled={isProcessing}
              className="px-4 py-2 text-gray-700 bg-gray-100 rounded-md hover:bg-gray-200 disabled:opacity-50"
            >
              Cancel
            </button>
            {step === 2 && (
              <button
                onClick={handleImport}
                disabled={isProcessing || importPreview.errors.length > 0}
                className="px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 disabled:opacity-50"
              >
                {isProcessing ? 'Importing...' : 'Start Import'}
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
