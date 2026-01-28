import { useState, useCallback, useEffect, useMemo, useRef } from "react";
import {
  loadCsv,
  loadGoogleSheet,
  buildPreview,
  generateEmails,
  pickCsvFile,
  pickResumeFile,
  pickTemplateFiles,
  exportTemplates,
  validateLicense,
  createDefaultTemplate,
  createDefaultProfile,
  saveProfilesToStorage,
  loadProfilesFromStorage,
  saveActiveProfileName,
  loadActiveProfileName,
  saveLicenseKey,
  loadLicenseKey,
  generateTemplateId,
  type PreviewRow,
  type Profile,
  type DataLoadResult,
  type Template,
} from "./engine";

// ============================================================
// Toast Component
// ============================================================

interface Toast {
  message: string;
  type: "success" | "error" | "warning";
}

// ============================================================
// License Modal Component
// ============================================================

interface LicenseModalProps {
  isOpen: boolean;
  onClose: () => void;
  licenseKey: string;
  setLicenseKey: (key: string) => void;
  isLicensed: boolean;
  onValidate: () => void;
  validating: boolean;
}

function LicenseModal({ isOpen, onClose, licenseKey, setLicenseKey, isLicensed, onValidate, validating }: LicenseModalProps) {
  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>License</h2>
        <div className={`license-status ${isLicensed ? "licensed" : "unlicensed"}`}>
          {isLicensed ? "✓ Licensed" : "● Unlicensed"}
        </div>
        <input
          type="text"
          value={licenseKey}
          onChange={(e) => setLicenseKey(e.target.value)}
          placeholder="Enter your license key..."
          className="license-input"
          onKeyDown={(e) => e.key === "Enter" && onValidate()}
        />
        <div className="modal-buttons">
          <button onClick={onValidate} disabled={validating || !licenseKey.trim()}>
            {validating ? "Validating..." : "Validate"}
          </button>
          <button onClick={onClose} className="btn-secondary">Close</button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// Preview Popup Modal Component
// ============================================================

interface PreviewPopupModalProps {
  isOpen: boolean;
  onClose: () => void;
  previewRows: PreviewRow[];
  onlyRecipients: boolean;
  setOnlyRecipients: (value: boolean) => void;
  onRefresh: () => void;
  loading: boolean;
  hasData: boolean;
  templates: Template[];
  overrides: Record<string, string>;
  onOverride: (email: string, templateId: string) => void;
  eligibleCount: number;
  onGenerate: () => void;
}

function PreviewPopupModal({
  isOpen,
  onClose,
  previewRows,
  onlyRecipients,
  setOnlyRecipients,
  onRefresh,
  loading,
  hasData,
  templates,
  overrides,
  onOverride,
  eligibleCount,
  onGenerate,
}: PreviewPopupModalProps) {
  const [selectedRowIndex, setSelectedRowIndex] = useState<number | null>(null);
  const [selectedEmail, setSelectedEmail] = useState<string | null>(null);

  // Preserve selection by email when rows change (e.g. after override refresh)
  useEffect(() => {
    if (selectedEmail && previewRows.length > 0) {
      const newIndex = previewRows.findIndex(
        (r) => r.email.toLowerCase() === selectedEmail.toLowerCase()
      );
      setSelectedRowIndex(newIndex >= 0 ? newIndex : null);
    } else {
      setSelectedRowIndex(null);
    }
  }, [previewRows, selectedEmail]);

  if (!isOpen) return null;

  const selectedRow = selectedRowIndex !== null ? previewRows[selectedRowIndex] : null;

  return (
    <div className="modal-overlay preview-popup-overlay" onClick={onClose}>
      <div className="modal preview-popup-modal" onClick={(e) => e.stopPropagation()}>
        <div className="preview-popup-header">
          <h2>Preview Recipients</h2>
          <div className="preview-popup-controls">
            <label className="checkbox-label">
              <input
                type="checkbox"
                checked={onlyRecipients}
                onChange={(e) => setOnlyRecipients(e.target.checked)}
              />
              Recipients Only
            </label>
            <button onClick={onRefresh} disabled={loading || !hasData} className="btn-icon" title="Refresh Preview">
              ↻
            </button>
            <button onClick={onClose} className="btn-secondary">Close</button>
          </div>
        </div>

        <div className="preview-popup-body">
          <div className="preview-popup-table-container">
            {loading ? (
              <div className="loading">Loading...</div>
            ) : previewRows.length === 0 ? (
              <div className="loading">Load data and refresh preview</div>
            ) : (
              <table className="preview-popup-table">
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Email</th>
                    <th>Firm</th>
                    <th style={{ width: "220px" }}>Template</th>
                  </tr>
                </thead>
                <tbody>
                  {previewRows.map((row, idx) => (
                    <tr
                      key={idx}
                      className={idx === selectedRowIndex ? "selected" : ""}
                      onClick={() => { setSelectedRowIndex(idx); setSelectedEmail(row.email); }}
                    >
                      <td>{row.name}</td>
                      <td>{row.email}</td>
                      <td>{row.firm}</td>
                      <td onClick={(e) => e.stopPropagation()}>
                        <select
                          className="table-select"
                          value={overrides[row.email.toLowerCase()] || (row.template_id || "")}
                          onChange={(e) => {
                            onOverride(row.email, e.target.value);
                            setSelectedRowIndex(idx);
                          }}
                        >
                          <option value="">-- Auto ({row.is_manual ? "Manual" : row.template_name}) --</option>
                          {templates.map((tpl) => (
                            <option key={tpl.id} value={tpl.id}>
                              {tpl.name}
                            </option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          <div className="preview-popup-inspector">
            <div className="section-title">Inspector</div>
            {selectedRow ? (() => {
              // Derive the actual template name from overrides if present
              const overrideId = overrides[selectedRow.email.toLowerCase()];
              const displayTemplate = overrideId
                ? templates.find((t) => t.id === overrideId)?.name || "(Unknown Override)"
                : selectedRow.template_name || "(None)";
              return (
                <div className="inspector-info">
                  <div><strong>Name:</strong> {selectedRow.name}</div>
                  <div><strong>Email:</strong> {selectedRow.email}</div>
                  <div><strong>Firm:</strong> {selectedRow.firm}</div>
                  <div><strong>Template:</strong> {displayTemplate}</div>
                  <div><strong>Source:</strong> {overrideId ? "Manual Override" : "Auto-Assigned"}</div>
                </div>
              );
            })() : (
              <div className="inspector-placeholder">
                Select a recipient to view details
              </div>
            )}
          </div>
        </div>

        <div className="preview-popup-footer">
          <div className="preview-footer-left">
            <span className="preview-popup-count">Visible: {previewRows.length} | Eligible: {eligibleCount}</span>
          </div>
          <button onClick={onGenerate} className="btn-primary">
            Generate {eligibleCount} Outlook Drafts
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// Confirm Modal Component
// ============================================================

interface ConfirmModalProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  message: string;
  onConfirm: () => void;
}

function ConfirmModal({ isOpen, onClose, title, message, onConfirm }: ConfirmModalProps) {
  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()} style={{ maxWidth: "400px" }}>
        <h2>{title}</h2>
        <p style={{ marginBottom: "1.5rem", lineHeight: "1.4" }}>{message}</p>
        <div className="modal-buttons">
          <button
            onClick={() => { onConfirm(); onClose(); }}
            className="btn-danger"
            autoFocus
          >
            Delete
          </button>
          <button onClick={onClose} className="btn-secondary">Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// Prompt Modal Component
// ============================================================

interface PromptModalProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  message: string;
  initialValue: string;
  onConfirm: (value: string) => void;
}

function PromptModal({ isOpen, onClose, title, message, initialValue, onConfirm }: PromptModalProps) {
  const [value, setValue] = useState(initialValue);

  // Update effect to reset value when modal opens key props change
  useEffect(() => {
    if (isOpen) setValue(initialValue);
  }, [isOpen, initialValue]);

  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()} style={{ maxWidth: "400px" }}>
        <h2>{title}</h2>
        <p style={{ marginBottom: "1rem" }}>{message}</p>
        <input
          type="text"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              onConfirm(value);
              onClose();
            }
          }}
          className="license-input"
          autoFocus
          style={{ marginBottom: "1.5rem" }}
        />
        <div className="modal-buttons">
          <button
            onClick={() => { onConfirm(value); onClose(); }}
            className="btn-primary"
            disabled={!value.trim()}
          >
            Save
          </button>
          <button onClick={onClose} className="btn-secondary">Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// Main App Component
// ============================================================

function App() {
  // ----------------------------------------
  // License State
  // ----------------------------------------
  const [licenseKey, setLicenseKey] = useState<string>(() => loadLicenseKey());
  const [isLicensed, setIsLicensed] = useState<boolean>(false);
  const [licenseModalOpen, setLicenseModalOpen] = useState<boolean>(false);
  const [validatingLicense, setValidatingLicense] = useState<boolean>(false);

  // ----------------------------------------
  // Profile State
  // ----------------------------------------
  const [profiles, setProfiles] = useState<Profile[]>(() => loadProfilesFromStorage());
  const [activeProfileName, setActiveProfileName] = useState<string>(() => loadActiveProfileName());

  const activeProfile = useMemo(() => {
    return profiles.find((p) => p.name === activeProfileName) || profiles[0];
  }, [profiles, activeProfileName]);

  // ----------------------------------------
  // Data State (derived from profile)
  // ----------------------------------------
  const [loadedData, setLoadedData] = useState<DataLoadResult | null>(null);

  // ----------------------------------------
  // Template Editor State
  // ----------------------------------------
  const [selectedTemplateIndex, setSelectedTemplateIndex] = useState<number>(0);
  const [editorText, setEditorText] = useState<string>("");
  const [editorName, setEditorName] = useState<string>("");
  const [editorManualOnly, setEditorManualOnly] = useState<boolean>(false);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState<boolean>(false);

  // ----------------------------------------
  // Preview State
  // ----------------------------------------
  const [previewRows, setPreviewRows] = useState<PreviewRow[]>([]);
  const [onlyRecipients, setOnlyRecipients] = useState<boolean>(true);


  // ----------------------------------------
  // UI State
  // ----------------------------------------
  const [loading, setLoading] = useState<boolean>(false);
  const [toast, setToast] = useState<Toast | null>(null);
  const [previewPopupOpen, setPreviewPopupOpen] = useState<boolean>(false);

  // Confirm Modal state
  const [confirmModalOpen, setConfirmModalOpen] = useState<boolean>(false);
  const [confirmConfig, setConfirmConfig] = useState<{
    title: string;
    message: string;
    onConfirm: () => void;
  } | null>(null);

  // Prompt Modal state
  const [promptModalOpen, setPromptModalOpen] = useState<boolean>(false);
  const [promptConfig, setPromptConfig] = useState<{
    title: string;
    message: string;
    initialValue: string;
    onConfirm: (value: string) => void;
  } | null>(null);

  // ----------------------------------------
  // Persistence Effects
  // ----------------------------------------
  useEffect(() => {
    saveProfilesToStorage(profiles);
  }, [profiles]);

  useEffect(() => {
    saveActiveProfileName(activeProfileName);
  }, [activeProfileName]);

  useEffect(() => {
    saveLicenseKey(licenseKey);
  }, [licenseKey]);

  // Sync editor state when template selection changes
  useEffect(() => {
    const templates = activeProfile?.templates || [];
    if (selectedTemplateIndex >= 0 && selectedTemplateIndex < templates.length) {
      const tpl = templates[selectedTemplateIndex];
      setEditorText(tpl.text);
      setEditorName(tpl.name);
      setEditorManualOnly(tpl.manual_only || false);
      setHasUnsavedChanges(false);
    }
  }, [selectedTemplateIndex, activeProfile?.templates]);

  // Validate license on startup
  useEffect(() => {
    if (licenseKey) {
      handleValidateLicense(true);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);





  // Track if we need to auto-load data (start true for initial load)
  const pendingAutoLoad = useRef<boolean>(true);

  // When profile name changes, mark that we need to load data
  useEffect(() => {
    pendingAutoLoad.current = true;
  }, [activeProfileName]);

  // ----------------------------------------
  // Keyboard Shortcuts
  // ----------------------------------------
  const saveHandlerRef = useRef<() => void>(() => { });

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key === "s") {
        e.preventDefault();
        saveHandlerRef.current();
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);

  // ----------------------------------------
  // Helpers
  // ----------------------------------------
  const getRowValue = (row: Record<string, string>, candidates: string[]): string => {
    const keys = Object.keys(row);
    for (const cand of candidates) {
      const match = keys.find((k) => k.toLowerCase() === cand.toLowerCase());
      if (match && row[match]) return row[match];
    }
    return "";
  };

  const showToast = useCallback((message: string, type: Toast["type"]) => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3000);
  }, []);

  const updateProfile = useCallback((updates: Partial<Profile>) => {
    setProfiles((prev) =>
      prev.map((p) => (p.name === activeProfileName ? { ...p, ...updates } : p))
    );
  }, [activeProfileName]);

  const selectedTemplate = useMemo(() => {
    const templates = activeProfile?.templates || [];
    return templates[selectedTemplateIndex] || null;
  }, [activeProfile?.templates, selectedTemplateIndex]);



  // ----------------------------------------
  // License Actions
  // ----------------------------------------
  const handleValidateLicense = useCallback(async (silent: boolean = false) => {
    if (!licenseKey.trim()) {
      if (!silent) showToast("Please enter a license key", "warning");
      return;
    }

    setValidatingLicense(true);
    try {
      const result = await validateLicense(licenseKey);
      if (result.success && result.data?.valid) {
        setIsLicensed(true);
        if (!silent) showToast(result.data.message || "License validated", "success");
      } else {
        setIsLicensed(false);
        if (!silent) showToast(result.error || result.data?.message || "Invalid license", "error");
      }
    } catch (error) {
      setIsLicensed(false);
      if (!silent) showToast("Failed to validate license", "error");
    } finally {
      setValidatingLicense(false);
    }
  }, [licenseKey, showToast]);

  // ----------------------------------------
  // Profile Actions
  // ----------------------------------------
  const handleNewProfile = useCallback(() => {
    const baseName = "New Profile";
    let name = baseName;
    let counter = 1;
    while (profiles.some((p) => p.name === name)) {
      name = `${baseName} ${counter++}`;
    }
    const newProfile = createDefaultProfile(name);
    setProfiles((prev) => [...prev, newProfile]);
    setActiveProfileName(name);
    setLoadedData(null);
    setPreviewRows([]);
  }, [profiles]);

  const handleDuplicateProfile = useCallback(() => {
    const newName = `${activeProfile.name} (Copy)`;
    let name = newName;
    let counter = 1;
    while (profiles.some((p) => p.name === name)) {
      name = `${activeProfile.name} (Copy ${counter++})`;
    }
    const duplicated: Profile = { ...activeProfile, name };
    setProfiles((prev) => [...prev, duplicated]);
    setActiveProfileName(name);
  }, [activeProfile, profiles]);

  const handleProfileNameChange = useCallback((newName: string) => {
    const currentName = activeProfileName;
    const trimmedName = newName.trim();

    // If empty or same name, just reset
    if (!trimmedName || trimmedName === currentName) {
      return;
    }

    if (profiles.some((p) => p.name === trimmedName && p.name !== currentName)) {
      showToast("A profile with that name already exists", "warning");
      return;
    }

    // Update profiles first
    setProfiles((prevProfiles) => {
      const updated = prevProfiles.map((p) =>
        p.name === currentName ? { ...p, name: trimmedName } : p
      );
      // After profiles update, update active name
      queueMicrotask(() => {
        setActiveProfileName(trimmedName);
      });
      return updated;
    });
    showToast(`Renamed to "${trimmedName}"`, "success");
  }, [activeProfileName, profiles, showToast]);

  const handleDeleteProfile = useCallback(() => {
    if (profiles.length <= 1) {
      showToast("Cannot delete the only profile", "warning");
      return;
    }

    setConfirmConfig({
      title: "Delete Profile",
      message: `Are you sure you want to delete profile "${activeProfile.name}"? This action cannot be undone.`,
      onConfirm: () => {
        const newProfiles = profiles.filter((p) => p.name !== activeProfileName);
        setProfiles(newProfiles);
        setActiveProfileName(newProfiles[0].name);
        setLoadedData(null);
        setPreviewRows([]);
      }
    });
    setConfirmModalOpen(true);
  }, [profiles, activeProfileName, activeProfile.name, showToast]);

  // ----------------------------------------
  // Data Loading Actions
  // ----------------------------------------
  const handleLoadData = useCallback(async (silent: boolean = false) => {
    // Validate that we have a data source configured
    const dataSourcePath = activeProfile.dataSource === "csv"
      ? activeProfile.csvPath
      : activeProfile.sheetUrl;

    if (!dataSourcePath || dataSourcePath.trim() === "") {
      if (!silent) showToast("No data source configured for this profile", "warning");
      return null;
    }

    setPreviewRows([]);
    setLoading(true);

    try {
      const result =
        activeProfile.dataSource === "csv"
          ? await loadCsv(activeProfile.csvPath)
          : await loadGoogleSheet(activeProfile.sheetUrl);

      if (result.success && result.data) {
        // Filter out rows where ALL columns are empty (completely blank rows)
        const filteredRows = result.data.rows.filter((row) => {
          // Keep the row if at least one column has a non-empty value (IGNORING the 'Generate' checkbox column)
          return result.data!.headers.some((header) => {
            if (header.toLowerCase() === "generate") return false;
            const value = row[header];
            return value !== undefined && value !== null && value.toString().trim() !== "";
          });
        });

        const filteredData = {
          ...result.data,
          rows: filteredRows,
          count: filteredRows.length,
        };

        setLoadedData(filteredData);
        if (!silent) {
          showToast(`Loaded ${filteredData.count} rows (${result.data.count - filteredData.count} empty rows filtered)`, "success");
        }

        // SMART PREVIEW GENERATION:
        // 1. Identify "eligible" rows (valid email + 'Generate' checked if present)
        // 2. Send ONLY eligible rows to valid generator
        // 3. Merge results back: Ineligible rows get blank template

        const allRows = filteredData.rows;
        const eligibleRows: Record<string, string>[] = [];
        const originalIndices: number[] = [];

        allRows.forEach((row, idx) => {
          // Check eligibility logic (same as Python basically, but we do it loosely here)
          // We need to send it to the engine to be sure? Actually engine does validation.
          // BUT if we send ineligible rows, they consume rotation slots.
          // So we filter by: has 'email', and if 'generate' col exists, it must be true-ish.

          let isEligible = true;
          if (!row["email"] && !row["Email"]) isEligible = false;

          // Check Generate column if it exists (case insensitive)
          const genKey = Object.keys(row).find(k => k.toLowerCase() === "generate");
          if (genKey) {
            const val = row[genKey]?.toString().toLowerCase();
            if (val === "false" || val === "0" || val === "no" || val === "") isEligible = false;
          }

          if (isEligible) {
            eligibleRows.push(row);
            originalIndices.push(idx);
          }
        });

        const previewResult = await buildPreview(
          { rows: eligibleRows, headers: filteredData.headers },
          activeProfile.templates,
          activeProfile.overrides,
          onlyRecipients
        );

        if (previewResult.success && previewResult.data) {
          // Reconstruct full list
          // If we sent all rows, we'd get all rows back.
          // Since we sent partial, we have to map back.

          // Actually, the engine preview returns what we sent.
          // We need to construct the FULL preview list for the UI.

          // If we are showing "recipients only" (onlyRecipients=true), we just show the result we got?
          // BUT the user wants "when all the rows are shown... ineligible... pure aesthetic".

          // So if onlyRecipients is TRUE, we just show eligible rows (as returned).
          // If onlyRecipients is FALSE (Show All), we need to merge.

          if (onlyRecipients) {
            setPreviewRows(previewResult.data.preview_rows);
          } else {
            // We need to merge eligible results with ineligible placeholders
            const fullPreview: PreviewRow[] = allRows.map((row, idx) => {
              const eligibleIdx = originalIndices.indexOf(idx);
              if (eligibleIdx !== -1) {
                // This was an eligible row, grab from result
                return previewResult.data!.preview_rows[eligibleIdx];
              } else {
                // Ineligible - return dummy with robust column search
                return {
                  name: getRowValue(row, ["name", "full name", "recipient", "candidate", "first name"]),
                  email: getRowValue(row, ["email", "e-mail", "mail", "address"]),
                  firm: getRowValue(row, ["firm", "company", "organization", "bank"]),
                  template_name: "",
                  template_id: null,
                  is_manual: false,
                  is_eligible: false
                } as PreviewRow;
              }
            });
            setPreviewRows(fullPreview);
          }
        }

        return filteredData;
      } else {
        setLoadedData(null);
        if (!silent) showToast(result.error || "Failed to load data", "error");
        return null;
      }
    } catch (error) {
      setLoadedData(null);
      if (!silent) showToast("Failed to load data", "error");
      return null;
    } finally {
      setLoading(false);
    }
  }, [activeProfile, onlyRecipients, showToast]);




  // When handleLoadData updates (which means activeProfile has updated),
  // check if we have a pending load
  useEffect(() => {
    if (pendingAutoLoad.current) {
      pendingAutoLoad.current = false;
      handleLoadData(true);
    }
  }, [handleLoadData]);

  const handlePickCsv = useCallback(async () => {
    const path = await pickCsvFile();
    if (path) {
      updateProfile({ csvPath: path });
    }
  }, [updateProfile]);

  const handlePickResume = useCallback(async () => {
    const path = await pickResumeFile();
    if (path) {
      updateProfile({ resumePath: path });
    }
  }, [updateProfile]);

  const handleOpenSheetInBrowser = useCallback(async () => {
    if (activeProfile.sheetUrl) {
      try {
        const { open } = await import('@tauri-apps/plugin-shell');
        await open(activeProfile.sheetUrl);
      } catch {
        // Fallback for browser testing
        window.open(activeProfile.sheetUrl, "_blank");
      }
    }
  }, [activeProfile.sheetUrl]);

  // ----------------------------------------
  // Template Actions
  // ----------------------------------------
  const handleCreateTemplate = useCallback(() => {
    const newTemplate = createDefaultTemplate();
    const templates = [...activeProfile.templates, newTemplate];
    updateProfile({ templates });
    setSelectedTemplateIndex(templates.length - 1);
  }, [activeProfile.templates, updateProfile]);

  const handleDeleteTemplate = useCallback(() => {
    if (activeProfile.templates.length <= 1) {
      showToast("Cannot delete the only template", "warning");
      return;
    }

    if (!selectedTemplate) return;

    setConfirmConfig({
      title: "Delete Template",
      message: `Are you sure you want to delete template "${selectedTemplate.name}"?`,
      onConfirm: () => {
        const templates = activeProfile.templates.filter((_, i) => i !== selectedTemplateIndex);

        // Setup cleanup of overrides for this deleted template
        const deletedTemplateId = activeProfile.templates[selectedTemplateIndex].id;
        const newOverrides = { ...activeProfile.overrides };

        // Remove any overrides that point to the deleted template
        let cleanedCount = 0;
        for (const [email, templateId] of Object.entries(newOverrides)) {
          if (templateId === deletedTemplateId) {
            delete newOverrides[email];
            cleanedCount++;
          }
        }

        updateProfile({ templates, overrides: newOverrides });
        setSelectedTemplateIndex(Math.max(0, selectedTemplateIndex - 1));

        if (cleanedCount > 0) {
          showToast(`Deleted template and cleared ${cleanedCount} override(s)`, "success");
        }
      }
    });
    setConfirmModalOpen(true);
  }, [activeProfile.templates, selectedTemplate, selectedTemplateIndex, updateProfile, showToast]);

  const handleMoveTemplateUp = useCallback(() => {
    if (selectedTemplateIndex <= 0) return;
    const templates = [...activeProfile.templates];
    [templates[selectedTemplateIndex - 1], templates[selectedTemplateIndex]] =
      [templates[selectedTemplateIndex], templates[selectedTemplateIndex - 1]];
    updateProfile({ templates });
    setSelectedTemplateIndex(selectedTemplateIndex - 1);
  }, [activeProfile.templates, selectedTemplateIndex, updateProfile]);

  const handleMoveTemplateDown = useCallback(() => {
    if (selectedTemplateIndex >= activeProfile.templates.length - 1) return;
    const templates = [...activeProfile.templates];
    [templates[selectedTemplateIndex], templates[selectedTemplateIndex + 1]] =
      [templates[selectedTemplateIndex + 1], templates[selectedTemplateIndex]];
    updateProfile({ templates });
    setSelectedTemplateIndex(selectedTemplateIndex + 1);
  }, [activeProfile.templates, selectedTemplateIndex, updateProfile]);

  const handleSaveTemplate = useCallback(() => {
    if (!selectedTemplate) return;
    const templates = activeProfile.templates.map((t, i) =>
      i === selectedTemplateIndex
        ? { ...t, name: editorName, text: editorText, manual_only: editorManualOnly }
        : t
    );
    updateProfile({ templates });
    setHasUnsavedChanges(false);
    showToast("Template saved", "success");
  }, [activeProfile.templates, selectedTemplateIndex, editorName, editorText, editorManualOnly, updateProfile, showToast, selectedTemplate]);

  useEffect(() => {
    saveHandlerRef.current = handleSaveTemplate;
  }, [handleSaveTemplate]);

  const handleRevertTemplate = useCallback(() => {
    if (selectedTemplate) {
      setEditorText(selectedTemplate.text);
      setEditorName(selectedTemplate.name);
      setEditorManualOnly(selectedTemplate.manual_only || false);
      setHasUnsavedChanges(false);
    }
  }, [selectedTemplate]);

  // ----------------------------------------
  // Template Import/Export Actions
  // ----------------------------------------
  const handleImportTemplates = useCallback(async () => {
    const files = await pickTemplateFiles();
    if (!files || files.length === 0) return;

    setLoading(true);
    try {
      let added = 0;
      const newTemplates = [...activeProfile.templates];

      for (const file of files) {
        newTemplates.push({
          id: generateTemplateId(),
          name: file.name,
          text: file.content,
          manual_only: false,
        });
        added++;
      }

      if (added > 0) {
        updateProfile({ templates: newTemplates });
        setSelectedTemplateIndex(newTemplates.length - 1);
        showToast(`Imported ${added} template(s)`, "success");
      }
    } catch (error) {
      showToast("Failed to import templates", "error");
    } finally {
      setLoading(false);
    }
  }, [activeProfile.templates, updateProfile, showToast]);

  const handleExportTemplates = useCallback(async () => {
    if (activeProfile.templates.length === 0) {
      showToast("No templates to export", "warning");
      return;
    }

    setLoading(true);
    try {
      const result = await exportTemplates(activeProfile.templates);
      if (result.success) {
        showToast(result.message || "Exported templates to Downloads", "success");
      } else {
        showToast(result.error || "Export failed", "error");
      }
    } catch (error) {
      showToast("Failed to export templates", "error");
    } finally {
      setLoading(false);
    }
  }, [activeProfile.templates, showToast]);

  // ----------------------------------------
  // Preview Actions
  // ----------------------------------------
  const handleRefreshPreview = useCallback(async () => {
    if (!loadedData) {
      showToast("Please load data first", "warning");
      return;
    }

    setLoading(true);
    try {
      // Filter eligible rows for rotation integrity
      const allRows = loadedData.rows;
      const eligibleRows: Record<string, string>[] = [];
      const originalIndices: number[] = [];

      allRows.forEach((row, idx) => {
        let isEligible = true;
        // Basic email check
        if (!row["email"] && !row["Email"]) isEligible = false;

        // Generate column check
        const genKey = Object.keys(row).find(k => k.toLowerCase() === "generate");
        if (genKey) {
          const val = row[genKey]?.toString().toLowerCase();
          if (val === "false" || val === "0" || val === "no" || val === "") isEligible = false;
        }

        if (isEligible) {
          eligibleRows.push(row);
          originalIndices.push(idx);
        }
      });

      const result = await buildPreview(
        { rows: eligibleRows, headers: loadedData.headers },
        activeProfile.templates,
        activeProfile.overrides,
        false // We ask engine for "all" (which is just the filtered set here)
      );

      if (result.success && result.data) {
        // Reconstruct full list depending on view mode
        // Note: The 'onlyRecipients' flag in buildPreview argument effectively is handled by our manual filtering.
        // We ALWAYS passed 'false' to buildPreview here in original code? 
        // No, original was `false` (meaning show all?). Wait.
        // buildPreview(..., onlyRecipients)
        // If we want "All Rows" in UI, we need to merge.

        // Actually, handleRefreshPreview is called by the UI button.
        // The modal state `onlyRecipients` (passed as prop to modal, but not accessed here directly? Ah `onlyRecipients` is in scope!)

        let finalRows: PreviewRow[] = [];

        if (onlyRecipients) {
          // If checking "Recipients Only", we just show the eligible ones we generated
          finalRows = result.data.preview_rows;
        } else {
          // Merge
          finalRows = allRows.map((row, idx) => {
            const eligibleIdx = originalIndices.indexOf(idx);
            if (eligibleIdx !== -1) {
              return result.data!.preview_rows[eligibleIdx];
            } else {
              return {
                name: getRowValue(row, ["name", "full name", "recipient", "candidate", "first name"]),
                email: getRowValue(row, ["email", "e-mail", "mail", "address"]),
                firm: getRowValue(row, ["firm", "company", "organization", "bank"]),
                template_name: "",
                template_id: null,
                is_manual: false,
                is_eligible: false
              } as PreviewRow;
            }
          });
        }

        setPreviewRows(finalRows);
        showToast(`Preview: ${finalRows.length} rows loaded`, "success");
      } else {
        showToast(result.error || "Failed to build preview", "error");
      }
    } catch (error) {
      showToast("Failed to build preview", "error");
    } finally {
      setLoading(false);
    }
  }, [loadedData, activeProfile.templates, activeProfile.overrides, showToast]);

  // ----------------------------------------
  // Manual Override Actions
  // ----------------------------------------
  const handleOverrideTemplate = useCallback((email: string, templateId: string) => {
    // Normalize email to lowercase to match engine lookup
    const normalizedEmail = email.toLowerCase().trim();
    const overrides = { ...activeProfile.overrides };

    if (templateId) {
      // Set the override to the selected template
      overrides[normalizedEmail] = templateId;
    } else {
      // If "-- Auto --" selected (empty value), delete the override
      delete overrides[normalizedEmail];
    }

    updateProfile({ overrides });
    // We don't automatically refresh preview here to allow batch updates if needed, 
    // but for UI responsiveness we probably should. Or we can let the caller handle it.
    // But `handleRefreshPreview` is stable, so calling it is fine.
    handleRefreshPreview();
  }, [activeProfile.overrides, updateProfile, handleRefreshPreview]);





  // ----------------------------------------
  // Generate Actions
  // ----------------------------------------
  const handleGenerate = useCallback(async () => {
    // Check for license key first
    if (!licenseKey.trim()) {
      showToast("Please enter a valid license key first", "warning");
      setLicenseModalOpen(true);
      return;
    }

    // Re-validate license before generating (matches legacy behavior)
    setLoading(true);
    try {
      const licenseResult = await validateLicense(licenseKey);
      const licenseValid = !!(licenseResult.success && licenseResult.data?.valid);
      if (licenseValid) {
        setIsLicensed(true);
      } else {
        setIsLicensed(false);
        showToast(licenseResult.error || licenseResult.data?.message || "License invalid or expired", "error");
        setLicenseModalOpen(true);
        setLoading(false);
        return;
      }
    } catch {
      setIsLicensed(false);
      showToast("Failed to validate license", "error");
      setLicenseModalOpen(true);
      setLoading(false);
      return;
    }

    // Refresh data silently before generating
    // This returns the fresh data to use immediately
    const freshData = await handleLoadData(true);
    if (!freshData) {
      // Error toast already handled by handleLoadData if not silent? 
      // We passed true (silent), so we should show error here if needed or handleLoadData shows *errors* even when silent?
      // Our implementation shows errors if !silent.
      // Wait, snippet says: "You only see something if there is an error."
      // So I changed handleLoadData to ONLY show success toast if !silent.
      // But errors? I kept "if (!silent) showToast(..., error)".
      // I should FIX handleLoadData to ALWAYS show errors.
      return;
    }

    if (freshData.rows.length === 0) {
      showToast("No recipients found in data", "warning");
      setLoading(false);
      return;
    }

    try {
      // Clean overrides: filter empty values and normalize email keys to lowercase
      const cleanOverrides: Record<string, string> = {};
      for (const [email, templateId] of Object.entries(activeProfile.overrides)) {
        if (templateId && templateId.trim() !== "") {
          cleanOverrides[email.toLowerCase().trim()] = templateId;
        }
      }

      console.log("=== GENERATE DEBUG ===");
      console.log("Rows:", freshData.rows.length);
      console.log("Headers:", freshData.headers);
      console.log("Templates:", activeProfile.templates);
      console.log("Original overrides:", activeProfile.overrides);
      console.log("Clean overrides:", cleanOverrides);
      console.log("Subject:", activeProfile.subjectTemplate);

      const result = await generateEmails(
        { rows: freshData.rows, headers: freshData.headers },
        activeProfile.templates,
        cleanOverrides,
        activeProfile.subjectTemplate,
        activeProfile.resumePath || undefined
      );

      console.log("Generate result:", result);

      if (result.success && result.data) {
        showToast(`Created ${result.data.created} Outlook drafts`, "success");
      } else {
        showToast(result.error || "Failed to generate emails", "error");
      }
    } catch (error) {
      console.error("Generate error:", error);
      showToast("Failed to generate emails", "error");
    } finally {
      setLoading(false);
    }
  }, [licenseKey, activeProfile, handleLoadData, showToast]);

  // ----------------------------------------
  // Render
  // ----------------------------------------
  // ----------------------------------------
  // Render
  // ----------------------------------------
  const visibleRows = useMemo(() => {
    if (onlyRecipients) {
      return previewRows.filter((r) => r.is_eligible);
    }
    return previewRows;
  }, [previewRows, onlyRecipients]);

  const eligibleCount = useMemo(() => {
    return previewRows.filter((r) => r.is_eligible).length;
  }, [previewRows]);

  return (
    <div className="app">
      {/* Header */}
      <header className="header">
        <div className="header-left">
          <h1>DraftMate</h1>
        </div>
        <div className="header-right">
          <button
            className="btn-primary"
            onClick={async () => {
              // Quiet refresh before opening
              await handleLoadData(true);
              setPreviewPopupOpen(true);
            }}
            style={{ marginRight: "1rem" }}
          >
            Preview Recipients
          </button>
          <div className="profile-controls">
            <span className="profile-label">Profile:</span>
            <select
              value={activeProfileName}
              onChange={(e) => {
                setActiveProfileName(e.target.value);
                setLoadedData(null);
                setPreviewRows([]);

                setSelectedTemplateIndex(0);
              }}
              className="profile-select"
            >
              {profiles.map((p) => (
                <option key={p.name} value={p.name}>
                  {p.name}
                </option>
              ))}
            </select>
            <div className="profile-buttons">
              <button
                onClick={() => {
                  setPromptConfig({
                    title: "Rename Profile",
                    message: "Enter a new name for this profile:",
                    initialValue: activeProfileName,
                    onConfirm: (newName) => handleProfileNameChange(newName)
                  });
                  setPromptModalOpen(true);
                }}
                className="btn-small"
                title="Rename Profile"
              >
                ✎
              </button>
              <button onClick={handleNewProfile} className="btn-small" title="New Profile">+</button>
              <button onClick={handleDuplicateProfile} className="btn-small" title="Duplicate Profile">⧉</button>
              <button onClick={handleDeleteProfile} className="btn-small btn-danger" title="Delete Profile">×</button>
            </div>
          </div>
          <button
            onClick={() => setLicenseModalOpen(true)}
            className={`license-btn ${isLicensed ? "licensed" : "unlicensed"}`}
          >
            {isLicensed ? "✓ Licensed" : "Enter License"}
          </button>
        </div>
      </header>



      <div className="main-content">
        {/* Left Sidebar: Data Source & Templates */}
        <aside className="sidebar">
          {/* Data Source Section */}
          <div className="section">
            <div className="section-title">Data Source</div>
            <div className="input-group">
              <select
                value={activeProfile.dataSource}
                onChange={(e) => updateProfile({ dataSource: e.target.value as "csv" | "sheet" })}
              >
                <option value="sheet">Google Sheet</option>
                <option value="csv">CSV File</option>
              </select>
            </div>

            {activeProfile.dataSource === "sheet" ? (
              <div className="input-group">
                <label>Google Sheet URL</label>
                <input
                  type="text"
                  value={activeProfile.sheetUrl}
                  onChange={(e) => updateProfile({ sheetUrl: e.target.value })}
                  placeholder="Paste Google Sheets URL..."
                />
                <div className="btn-row">
                  <button onClick={() => handleLoadData(false)} disabled={loading || !activeProfile.sheetUrl}>
                    Load
                  </button>
                  <button onClick={handleOpenSheetInBrowser} disabled={!activeProfile.sheetUrl} className="btn-secondary">
                    Open
                  </button>
                </div>
              </div>
            ) : (
              <div className="input-group">
                <label>CSV File</label>
                <div className="file-picker">
                  <input
                    type="text"
                    value={activeProfile.csvPath}
                    readOnly
                    placeholder="No file selected"
                  />
                  <button onClick={handlePickCsv} className="btn-secondary">Browse</button>
                </div>
                <button onClick={() => handleLoadData(false)} disabled={loading || !activeProfile.csvPath}>
                  Load Data
                </button>
              </div>
            )}
          </div>

          {/* Email Settings Section */}
          <div className="section">
            <div className="section-title">Email Settings</div>
            <div className="input-group">
              <label>Subject Template</label>
              <input
                type="text"
                value={activeProfile.subjectTemplate}
                onChange={(e) => updateProfile({ subjectTemplate: e.target.value })}
                placeholder="Enter subject line... (e.g. Duke Student interested in IB at {firm})"
              />
            </div>
            <div className="input-group">
              <label>Resume PDF</label>
              <div className="file-picker">
                <input
                  type="text"
                  value={activeProfile.resumePath}
                  readOnly
                  placeholder="No file selected"
                />
                <button onClick={handlePickResume} className="btn-secondary">Browse</button>
              </div>
            </div>
          </div>

          {/* Templates Section */}
          <div className="section templates-section">
            <div className="section-title">
              Templates ({activeProfile.templates.length})
              <div className="template-actions">
                <button onClick={handleCreateTemplate} className="btn-icon" title="New Template">+</button>
                <button onClick={handleDeleteTemplate} className="btn-icon" title="Delete Template">−</button>
                <button onClick={handleMoveTemplateUp} className="btn-icon" title="Move Up">↑</button>
                <button onClick={handleMoveTemplateDown} className="btn-icon" title="Move Down">↓</button>
              </div>
            </div>
            <div className="template-list">
              {activeProfile.templates.map((tpl, idx) => (
                <div
                  key={tpl.id}
                  className={`template-item ${idx === selectedTemplateIndex ? "selected" : ""} ${tpl.manual_only ? "manual-only" : ""}`}
                  onClick={() => setSelectedTemplateIndex(idx)}
                >
                  <span className="template-name">{tpl.name}</span>
                  {tpl.manual_only && <span className="badge">M</span>}
                </div>
              ))}
            </div>
            <div className="template-import-export">
              <button onClick={handleImportTemplates} className="btn-secondary" disabled={loading}>
                Import .txt
              </button>
              <button onClick={handleExportTemplates} className="btn-secondary" disabled={loading}>
                Export .zip
              </button>
            </div>
          </div>
        </aside>

        {/* Center: Template Editor */}
        <div className="editor-panel">
          <div className="editor-header">
            <div className="editor-title">
              {selectedTemplate ? (
                <>
                  <input
                    type="text"
                    value={editorName}
                    onChange={(e) => { setEditorName(e.target.value); setHasUnsavedChanges(true); }}
                    className="template-name-input"
                  />
                  {hasUnsavedChanges && <span className="unsaved-indicator">•</span>}
                </>
              ) : (
                "No Template Selected"
              )}
            </div>
            <div className="editor-actions">
              <label className="checkbox-label">
                <input
                  type="checkbox"
                  checked={editorManualOnly}
                  onChange={(e) => { setEditorManualOnly(e.target.checked); setHasUnsavedChanges(true); }}
                />
                Manual Only
              </label>
              <button onClick={handleSaveTemplate} disabled={!hasUnsavedChanges}>Save</button>
              <button onClick={handleRevertTemplate} disabled={!hasUnsavedChanges} className="btn-secondary">Revert</button>
            </div>
          </div>
          <textarea
            className="template-editor"
            value={editorText}
            onChange={(e) => { setEditorText(e.target.value); setHasUnsavedChanges(true); }}
            placeholder="Enter your email template here...&#10;&#10;Use placeholders like:&#10;{first name}, {last name}, {firm}, {school}"
            disabled={!selectedTemplate}
          />
          <div className="placeholder-hints">
            Placeholders: {"{first name}"}, {"{last name}"}, {"{full name}"}, {"{firm}"}, {"{school}"}, or any column header
          </div>
        </div>


      </div>

      {/* Main Screen Actions */}
      <div className="footer">
        <button
          onClick={handleGenerate}
          className="btn-primary btn-wide"
          disabled={loading || !loadedData || eligibleCount === 0}
        >
          Generate {eligibleCount > 0 ? `${eligibleCount} ` : ""}Outlook Drafts
        </button>
      </div>



      {/* License Modal */}
      <LicenseModal
        isOpen={licenseModalOpen}
        onClose={() => setLicenseModalOpen(false)}
        licenseKey={licenseKey}
        setLicenseKey={setLicenseKey}
        isLicensed={isLicensed}
        onValidate={() => handleValidateLicense(false)}
        validating={validatingLicense}
      />

      {/* Preview Popup Modal */}
      <PreviewPopupModal
        isOpen={previewPopupOpen}
        onClose={() => setPreviewPopupOpen(false)}
        previewRows={visibleRows}
        eligibleCount={eligibleCount}
        onGenerate={handleGenerate}
        onlyRecipients={onlyRecipients}
        setOnlyRecipients={setOnlyRecipients}
        onRefresh={handleRefreshPreview}
        loading={loading}
        hasData={!!loadedData}
        templates={activeProfile.templates}
        overrides={activeProfile.overrides}
        onOverride={handleOverrideTemplate}
      />

      {/* Confirm Modal */}
      <ConfirmModal
        isOpen={confirmModalOpen}
        onClose={() => setConfirmModalOpen(false)}
        title={confirmConfig?.title || ""}
        message={confirmConfig?.message || ""}
        onConfirm={confirmConfig?.onConfirm || (() => { })}
      />

      {/* Prompt Modal */}
      <PromptModal
        isOpen={promptModalOpen}
        onClose={() => setPromptModalOpen(false)}
        title={promptConfig?.title || ""}
        message={promptConfig?.message || ""}
        initialValue={promptConfig?.initialValue || ""}
        onConfirm={promptConfig?.onConfirm || (() => { })}
      />

      {/* Toast */}
      {toast && <div className={`toast ${toast.type}`}>{toast.message}</div>}
    </div >
  );
}

export default App;
