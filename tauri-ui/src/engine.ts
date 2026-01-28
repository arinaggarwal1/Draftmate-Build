/**
 * DraftMate Engine Bridge
 *
 * This module provides the interface between the Tauri frontend
 * and the Python engine via Rust commands.
 */

import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";

// ============================================================
// Types
// ============================================================

export interface EngineResponse<T> {
  success: boolean;
  data: T | null;
  error: string | null;
}

export interface DataLoadResult {
  rows: Record<string, string>[];
  headers: string[];
  count: number;
}

export interface PreviewRow {
  name: string;
  email: string;
  firm: string;
  template_name: string;
  template_id: string | null;
  is_manual: boolean;
  is_eligible: boolean;
}

export interface PreviewResult {
  preview_rows: PreviewRow[];
  count: number;
}

export interface GenerateResult {
  created: number;
}

export interface Template {
  id: string;
  name: string;
  text: string;
  manual_only?: boolean;
}

export interface Profile {
  name: string;
  dataSource: "csv" | "sheet";
  sheetUrl: string;
  csvPath: string;
  resumePath: string;
  subjectTemplate: string;
  templates: Template[];
  overrides: Record<string, string>;
}

// ============================================================
// Engine Command Runner
// ============================================================

/**
 * Run a Python engine command via Tauri invoke and return parsed JSON response.
 */
async function runEngineCommand<T>(args: string[]): Promise<EngineResponse<T>> {
  try {
    const stdout = await invoke<string>("run_engine", { args });
    const response = JSON.parse(stdout) as EngineResponse<T>;
    return response;
  } catch (error) {
    return {
      success: false,
      data: null,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

// ============================================================
// Data Loading
// ============================================================

export async function loadCsv(path: string): Promise<EngineResponse<DataLoadResult>> {
  return runEngineCommand<DataLoadResult>(["load-csv", path]);
}

export async function loadGoogleSheet(url: string): Promise<EngineResponse<DataLoadResult>> {
  return runEngineCommand<DataLoadResult>(["load-sheet", url]);
}

// ============================================================
// Preview
// ============================================================

export async function buildPreview(
  data: { rows: Record<string, string>[]; headers: string[] },
  templates: Template[],
  overrides: Record<string, string> = {},
  onlyRecipients: boolean = true
): Promise<EngineResponse<PreviewResult>> {
  const args = [
    "preview",
    "--data",
    JSON.stringify(data),
    "--templates",
    JSON.stringify(templates),
    "--overrides",
    JSON.stringify(overrides),
  ];

  if (!onlyRecipients) {
    args.push("--all-rows");
  }

  return runEngineCommand<PreviewResult>(args);
}

// ============================================================
// Generation
// ============================================================

export async function generateEmails(
  data: { rows: Record<string, string>[]; headers: string[] },
  templates: Template[],
  overrides: Record<string, string>,
  subjectTemplate: string,
  resumePath?: string,
  dryRun: boolean = false
): Promise<EngineResponse<GenerateResult>> {
  const args = [
    "generate",
    "--data",
    JSON.stringify(data),
    "--templates",
    JSON.stringify(templates),
    "--overrides",
    JSON.stringify(overrides),
    "--subject",
    subjectTemplate,
  ];

  if (resumePath) {
    args.push("--resume", resumePath);
  }

  if (dryRun) {
    args.push("--dry-run");
  }

  return runEngineCommand<GenerateResult>(args);
}

// ============================================================
// File Dialogs
// ============================================================

export async function pickCsvFile(): Promise<string | null> {
  try {
    const selected = await open({
      multiple: false,
      filters: [{ name: "CSV Files", extensions: ["csv"] }],
    });
    return selected as string | null;
  } catch {
    return null;
  }
}

export async function pickResumeFile(): Promise<string | null> {
  try {
    const selected = await open({
      multiple: false,
      filters: [{ name: "PDF Files", extensions: ["pdf"] }],
    });
    return selected as string | null;
  } catch {
    return null;
  }
}

export async function pickTemplateFile(): Promise<string | null> {
  try {
    const selected = await open({
      multiple: false,
      filters: [{ name: "Text Files", extensions: ["txt"] }],
    });
    return selected as string | null;
  } catch {
    return null;
  }
}

export async function pickExportLocation(defaultName: string): Promise<string | null> {
  try {
    const selected = await save({
      defaultPath: defaultName,
      filters: [{ name: "ZIP Files", extensions: ["zip"] }],
    });
    return selected as string | null;
  } catch {
    return null;
  }
}

export interface TemplateFile {
  name: string;
  content: string;
}

/**
 * Pick multiple .txt files for template import.
 * Returns array of { name, content } objects.
 */
export async function pickTemplateFiles(): Promise<TemplateFile[] | null> {
  try {
    const selected = await open({
      multiple: true,
      filters: [{ name: "Text Files", extensions: ["txt"] }],
    });

    if (!selected || (Array.isArray(selected) && selected.length === 0)) {
      return null;
    }

    const paths = Array.isArray(selected) ? selected : [selected];

    // Use engine to read the files
    const result = await runEngineCommand<{ files: TemplateFile[] }>(["read-files", ...paths]);
    if (result.success && result.data) {
      return result.data.files;
    }
    return null;
  } catch {
    return null;
  }
}

export interface ExportResult {
  success: boolean;
  message?: string;
  error?: string;
}

/**
 * Export templates to a ZIP file in the Downloads folder.
 */
export async function exportTemplates(templates: Template[]): Promise<ExportResult> {
  const result = await runEngineCommand<{ path: string }>(["export-templates", "--templates", JSON.stringify(templates)]);
  if (result.success && result.data) {
    return { success: true, message: `Exported to ${result.data.path}` };
  }
  return { success: false, error: result.error || "Export failed" };
}

// ============================================================
// Utilities
// ============================================================

export function generateTemplateId(): string {
  return `tpl_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

export function createDefaultTemplate(): Template {
  return {
    id: generateTemplateId(),
    name: "New Template",
    text: "Dear {first name},\n\n\n\nBest regards",
    manual_only: false,
  };
}

export function createDefaultProfile(name: string): Profile {
  return {
    name,
    dataSource: "sheet",
    sheetUrl: "",
    csvPath: "",
    resumePath: "",
    subjectTemplate: "{first name} - Networking Request",
    templates: [createDefaultTemplate()],
    overrides: {},
  };
}

// ============================================================
// Local Storage Persistence
// ============================================================

const STORAGE_KEY = "draftmate_profiles";
const ACTIVE_PROFILE_KEY = "draftmate_active_profile";
const LICENSE_KEY = "draftmate_license_key";

export function saveProfilesToStorage(profiles: Profile[]): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(profiles));
}

export function loadProfilesFromStorage(): Profile[] {
  try {
    const data = localStorage.getItem(STORAGE_KEY);
    if (data) {
      return JSON.parse(data);
    }
  } catch {
    // Ignore parse errors
  }
  return [createDefaultProfile("Default")];
}

export function saveActiveProfileName(name: string): void {
  localStorage.setItem(ACTIVE_PROFILE_KEY, name);
}

export function loadActiveProfileName(): string {
  return localStorage.getItem(ACTIVE_PROFILE_KEY) || "Default";
}

export function saveLicenseKey(key: string): void {
  localStorage.setItem(LICENSE_KEY, key);
}

export function loadLicenseKey(): string {
  return localStorage.getItem(LICENSE_KEY) || "";
}

// ============================================================
// License Validation
// ============================================================

export interface LicenseResult {
  valid: boolean;
  message: string;
}

export async function validateLicense(licenseKey: string): Promise<EngineResponse<LicenseResult>> {
  return runEngineCommand<LicenseResult>(["validate-license", licenseKey]);
}
