import { read, utils } from 'xlsx'

export type Material = {
  Record_ID: number
  CPSE: string
  Material_Code: string
  Material_Description: string
  Category: string
  Material_Type: string
  Material: string
  Grade: string
  Primary_Spec: string
  Secondary_Spec: string
  UOM: string
  Unit_Price_INR: number
  Annual_Quantity: number
  Procurement_Year: number
  [key: string]: string | number
}
export type GroundTruthPair = { Record_A: number; Record_B: number; Ground_Truth_Label: string; Reason: string }
export type MatchPair = { a: Material; b: Material; score: number; type: string; confidence: string; matching: string[]; conflicts: string[]; explanation: string; descriptionScore: number }
export type DecisionStatus = 'Approved' | 'Deferred' | 'Rejected'
export type MaterialData = { materials: Material[]; groundTruth: GroundTruthPair[]; pairs: MatchPair[]; stats: { total: number; cpse: number; categories: number; potential: number; equivalent: number; underReview: number; harmonization: number; tp: number; fp: number; fn: number; tn: number; precision: number; recall: number; f1: number }; filename?: string; hasGroundTruth?: boolean }

export const MATCH_WEIGHTS = { description: 0.3, category: 0.1, materialType: 0.1, material: 0.1, grade: 0.15, primarySpec: 0.15, secondarySpec: 0.05, uom: 0.05 }
export const normalize = (value: unknown) => String(value ?? '').toLowerCase().replace(/[µ²×]/g, (v) => ({ 'µ': 'u', '²': '2', '×': 'x' })[v] ?? v).replace(/[^a-z0-9]+/g, ' ').trim()
const aliases: Record<string, string> = { ss: 'stainless steel', cs: 'carbon steel', ms: 'mild steel', hex: 'hexagonal', dia: 'diameter', od: 'outer diameter', id: 'inner diameter', pcs: 'nos', pc: 'nos' }
export const normalized = (value: unknown) => Object.entries(aliases).reduce((text, [from, to]) => text.replace(new RegExp(`\\b${from}\\b`, 'g'), to), normalize(value))
const fieldLabels: Record<string, string> = { Category: 'Category', Material_Type: 'Material type', Material: 'Material', Grade: 'Grade', Primary_Spec: 'Primary specification', Secondary_Spec: 'Secondary specification', UOM: 'UOM' }
const fieldAliases: Record<keyof Material, string[]> = {
  Record_ID: ['record id', 'record_id', 'id'], CPSE: ['cpse', 'company', 'organization', 'organisation'], Material_Code: ['material code', 'material_code', 'code', 'item code', 'item_code', 'legacy code'], Material_Description: ['material description', 'description', 'item description'], Category: ['category', 'material category', 'item category'], Material_Type: ['material type', 'type'], Material: ['material', 'base material'], Grade: ['grade', 'material grade'], Primary_Spec: ['primary specification', 'primary specifications', 'specification', 'specifications', 'specs'], Secondary_Spec: ['secondary specification', 'secondary specifications'], UOM: ['uom', 'unit', 'unit of measure'], Unit_Price_INR: ['unit price', 'unit price inr', 'price', 'rate'], Annual_Quantity: ['annual quantity', 'quantity', 'annual qty', 'procurement quantity'], Procurement_Year: ['procurement year', 'year'],
}
const canonical = (key: string) => normalize(key).replace(/\s+/g, ' ')
function mapRow(row: Record<string, unknown>, index: number): Material {
  const keys = Object.keys(row)
  const value = (field: keyof Material) => { const found = keys.find((key) => fieldAliases[field].includes(canonical(key))); return found ? row[found] : '' }
  const text = (field: keyof Material, fallback = '') => String(value(field) ?? fallback).trim()
  return { Record_ID: Number(value('Record_ID')) || index + 1, CPSE: text('CPSE', 'Not specified'), Material_Code: text('Material_Code', `MAT-${index + 1}`), Material_Description: text('Material_Description', text('Material_Code', `Material ${index + 1}`)), Category: text('Category', 'Not specified'), Material_Type: text('Material_Type'), Material: text('Material'), Grade: text('Grade'), Primary_Spec: text('Primary_Spec'), Secondary_Spec: text('Secondary_Spec'), UOM: text('UOM'), Unit_Price_INR: Number(value('Unit_Price_INR')) || 0, Annual_Quantity: Number(value('Annual_Quantity')) || 0, Procurement_Year: Number(value('Procurement_Year')) || 0 }
}
export function parseRows(rows: Record<string, unknown>[], filename = 'Uploaded dataset'): MaterialData {
  const materials = rows.map(mapRow).filter((m) => m.Material_Description || m.Material_Code)
  return buildMaterialData(materials, [], filename)
}
export function parseWorkbookBuffer(buffer: ArrayBuffer, filename: string): MaterialData {
  const workbook = read(buffer, { type: 'array', cellDates: true })
  const materialSheet = workbook.SheetNames.find((name) => normalize(name).includes('material')) || workbook.SheetNames[0]
  const rows = materialSheet ? utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[materialSheet], { defval: '' }) : []
  const headers = rows[0] ? Object.keys(rows[0]).map(canonical) : []
  const hasDescription = headers.some((header) => fieldAliases.Material_Description.some((alias) => canonical(alias) === header))
  const hasCode = headers.some((header) => fieldAliases.Material_Code.some((alias) => canonical(alias) === header))
  if (!rows.length) throw new Error('No material records found in the uploaded file.')
  if (!hasDescription && !hasCode) throw new Error('Could not identify a material description or material code column. Expected aliases include Description, Material Description, Item Description, Material Code, or Item Code.')
  const truthSheet = workbook.SheetNames.find((name) => normalize(name).includes('ground truth') || normalize(name).includes('evaluation'))
  const groundTruth = truthSheet ? utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[truthSheet], { defval: '' }).map((row) => ({ Record_A: Number(row.Record_A) || 0, Record_B: Number(row.Record_B) || 0, Ground_Truth_Label: String(row.Ground_Truth_Label ?? row.Label ?? ''), Reason: String(row.Reason ?? '') })) : []
  return buildMaterialData(rows.map(mapRow), groundTruth, filename)
}
export function scorePair(a: Material, b: Material): MatchPair {
  const fields: Array<keyof Material> = ['Category', 'Material_Type', 'Material', 'Grade', 'Primary_Spec', 'Secondary_Spec', 'UOM']
  const matching = fields.filter((field) => normalized(a[field]) && normalized(a[field]) === normalized(b[field])).map((field) => fieldLabels[field])
  const conflicts = fields.filter((field) => normalized(a[field]) && normalized(b[field]) && normalized(a[field]) !== normalized(b[field])).map((field) => fieldLabels[field])
  const tokensA = new Set(normalized(a.Material_Description).split(' ').filter(Boolean)); const tokensB = new Set(normalized(b.Material_Description).split(' ').filter(Boolean)); const union = new Set([...tokensA, ...tokensB]); const descriptionScore = union.size ? [...tokensA].filter((token) => tokensB.has(token)).length / union.size : 0
  const values: Record<string, number> = { description: descriptionScore, category: matching.includes('Category') ? 1 : 0, materialType: matching.includes('Material type') ? 1 : 0, material: matching.includes('Material') ? 1 : 0, grade: matching.includes('Grade') ? 1 : 0, primarySpec: matching.includes('Primary specification') ? 1 : 0, secondarySpec: matching.includes('Secondary specification') ? 1 : 0, uom: matching.includes('UOM') ? 1 : 0 }
  let raw = Object.entries(MATCH_WEIGHTS).reduce((sum, [key, weight]) => sum + values[key] * weight, 0)
  if (conflicts.includes('Grade') || conflicts.includes('Primary specification')) raw *= 0.6
  const score = Math.round(raw * 100); const type = score >= 85 ? 'Highly Similar' : score >= 65 ? 'Potential Equivalent' : score >= 45 ? 'Near Duplicate' : 'Low Similarity'; const confidence = score >= 75 && conflicts.length <= 1 ? 'Likely Equivalent' : score >= 45 ? 'Needs Review' : 'Low Confidence'
  return { a, b, score, type, confidence, matching, conflicts, descriptionScore: Math.round(descriptionScore * 100), explanation: `${matching.length} of ${fields.length} structured attributes align; description token overlap is ${Math.round(descriptionScore * 100)}%.${conflicts.length ? ` Conflicts: ${conflicts.join(', ')}.` : ''}` }
}
export function buildMaterialData(materials: Material[], groundTruth: GroundTruthPair[] = [], filename?: string): MaterialData {
  const pairs = materials
    .flatMap((a, index) => {
      return materials.slice(index + 1)
        .filter((b) => a.CPSE !== b.CPSE || new Set(materials.map((m) => m.CPSE)).size === 1)
        .map((b) => scorePair(a, b))
    })
    .filter((pair) => pair.score >= 35)
    .sort((a, b) => b.score - a.score)
  const truth = new Set(groundTruth.filter((pair) => ['EXACT_DUPLICATE', 'EQUIVALENT'].includes(pair.Ground_Truth_Label.toUpperCase())).map((pair) => [pair.Record_A, pair.Record_B].sort((x, y) => x - y).join('-'))); const predicted = new Set(pairs.filter((pair) => pair.score >= 60).map((pair) => [pair.a.Record_ID, pair.b.Record_ID].sort((x, y) => x - y).join('-'))); const tp = [...truth].filter((key) => predicted.has(key)).length; const fp = [...predicted].filter((key) => !truth.has(key)).length; const fn = [...truth].filter((key) => !predicted.has(key)).length; const tn = Math.max(0, materials.length * (materials.length - 1) / 2 - tp - fp - fn); const precision = tp + fp ? tp / (tp + fp) : 0; const recall = tp + fn ? tp / (tp + fn) : 0; const f1 = precision + recall ? 2 * precision * recall / (precision + recall) : 0
  return { materials, groundTruth, pairs, filename, hasGroundTruth: groundTruth.length > 0, stats: { total: materials.length, cpse: new Set(materials.map((m) => m.CPSE).filter(Boolean)).size, categories: new Set(materials.map((m) => m.Category).filter(Boolean)).size, potential: pairs.filter((p) => p.score >= 60).length, equivalent: 0, underReview: pairs.filter((p) => p.score >= 45 && p.score < 75).length, harmonization: 0, tp, fp, fn, tn, precision, recall, f1 } }
}
export const keyOf = (p: MatchPair) => `${p.a.Record_ID}-${p.b.Record_ID}`
