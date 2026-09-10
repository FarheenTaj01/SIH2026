import { readFile } from 'node:fs/promises'
import path from 'node:path'
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
}
export type GroundTruthPair = { Record_A: number; Record_B: number; Ground_Truth_Label: string; Reason: string }
export type MatchPair = { a: Material; b: Material; score: number; type: string; confidence: string; matching: string[]; conflicts: string[]; explanation: string; descriptionScore: number }

export const MATCH_WEIGHTS = { description: 0.3, category: 0.1, materialType: 0.1, material: 0.1, grade: 0.15, primarySpec: 0.15, secondarySpec: 0.05, uom: 0.05 }
export const normalize = (value: unknown) => String(value ?? '').toLowerCase().replace(/[µ²×]/g, (v) => ({ 'µ': 'u', '²': '2', '×': 'x' })[v] ?? v).replace(/[^a-z0-9]+/g, ' ').trim()
const aliases: Record<string, string> = { ss: 'stainless steel', cs: 'carbon steel', ms: 'mild steel', hex: 'hexagonal', dia: 'diameter', od: 'outer diameter', id: 'inner diameter', pcs: 'nos', pc: 'nos', '3 phase': 'three phase', '3-phase': 'three phase' }
export const normalized = (value: unknown) => Object.entries(aliases).reduce((text, [from, to]) => text.replace(new RegExp(`\\b${from}\\b`, 'g'), to), normalize(value))
const fieldLabels: Record<string, string> = { Category: 'Category', Material_Type: 'Material type', Material: 'Material', Grade: 'Grade', Primary_Spec: 'Primary specification', Secondary_Spec: 'Secondary specification', UOM: 'UOM' }

export function scorePair(a: Material, b: Material): MatchPair {
  const fields: Array<keyof Material> = ['Category', 'Material_Type', 'Material', 'Grade', 'Primary_Spec', 'Secondary_Spec', 'UOM']
  const matching = fields.filter((field) => normalized(a[field]) && normalized(a[field]) === normalized(b[field])).map((field) => fieldLabels[field])
  const conflicts = fields.filter((field) => normalized(a[field]) && normalized(b[field]) && normalized(a[field]) !== normalized(b[field])).map((field) => fieldLabels[field])
  const tokensA = new Set(normalized(a.Material_Description).split(' ').filter(Boolean)); const tokensB = new Set(normalized(b.Material_Description).split(' ').filter(Boolean)); const union = new Set([...tokensA, ...tokensB])
  const descriptionScore = union.size ? [...tokensA].filter((token) => tokensB.has(token)).length / union.size : 0
  const values: Record<string, number> = { description: descriptionScore, category: matching.includes('Category') ? 1 : 0, materialType: matching.includes('Material type') ? 1 : 0, material: matching.includes('Material') ? 1 : 0, grade: matching.includes('Grade') ? 1 : 0, primarySpec: matching.includes('Primary specification') ? 1 : 0, secondarySpec: matching.includes('Secondary specification') ? 1 : 0, uom: matching.includes('UOM') ? 1 : 0 }
  const score = Math.round(Object.entries(MATCH_WEIGHTS).reduce((sum, [key, weight]) => sum + values[key] * weight, 0) * 100)
  const type = score >= 85 ? 'Highly Similar' : score >= 65 ? 'Potential Equivalent' : score >= 45 ? 'Near Duplicate' : 'Low Similarity'
  const confidence = score >= 75 && conflicts.length <= 1 ? 'Likely Equivalent' : score >= 45 ? 'Needs Review' : 'Low Confidence'
  return { a, b, score, type, confidence, matching, conflicts, descriptionScore: Math.round(descriptionScore * 100), explanation: `${matching.length} of ${fields.length} structured attributes align; description token overlap is ${Math.round(descriptionScore * 100)}%.` }
}

export async function getMaterialData() {
  const file = await readFile(path.join(process.cwd(), 'data/SIH26099_50row_dataset_with_ground_truth_labels-59b7df.xlsx'))
  const workbook = read(file, { type: 'buffer', cellDates: true })
  const materials = utils.sheet_to_json<Material>(workbook.Sheets.Material_Data, { defval: '' }).map((row) => ({ ...row, Record_ID: Number(row.Record_ID), Unit_Price_INR: Number(row.Unit_Price_INR) || 0, Annual_Quantity: Number(row.Annual_Quantity) || 0, Procurement_Year: Number(row.Procurement_Year) || 0 }))
  const groundTruth = utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets.Ground_Truth_Pairs, { defval: '' }).map((row) => ({ Record_A: Number(row.Record_A) || 0, Record_B: Number(row.Record_B) || 0, Ground_Truth_Label: String(row.Ground_Truth_Label ?? ''), Reason: String(row.Reason ?? '') }))
  const pairs = materials.flatMap((a, index) => materials.slice(index + 1).filter((b) => a.CPSE !== b.CPSE).map((b) => scorePair(a, b))).filter((pair) => pair.score >= 35).sort((a, b) => b.score - a.score)
  const truth = new Set(groundTruth.filter((pair) => ['EXACT_DUPLICATE', 'EQUIVALENT'].includes(pair.Ground_Truth_Label)).map((pair) => [pair.Record_A, pair.Record_B].sort((x, y) => x - y).join('-')))
  const predicted = new Set(pairs.filter((pair) => pair.score >= 60).map((pair) => [pair.a.Record_ID, pair.b.Record_ID].sort((x, y) => x - y).join('-')))
  const totalPairs = materials.length * (materials.length - 1) / 2; const tp = [...truth].filter((key) => predicted.has(key)).length; const fp = [...predicted].filter((key) => !truth.has(key)).length; const fn = [...truth].filter((key) => !predicted.has(key)).length; const tn = Math.max(0, totalPairs - tp - fp - fn); const precision = tp + fp ? tp / (tp + fp) : 0; const recall = tp + fn ? tp / (tp + fn) : 0; const f1 = precision + recall ? 2 * precision * recall / (precision + recall) : 0
  return { materials, groundTruth, pairs, stats: { total: materials.length, cpse: new Set(materials.map((m) => m.CPSE)).size, categories: new Set(materials.map((m) => m.Category)).size, potential: pairs.filter((p) => p.score >= 60).length, equivalent: predicted.size, underReview: pairs.filter((p) => p.score >= 45 && p.score < 75).length, harmonization: Math.round((predicted.size / Math.max(1, materials.length)) * 100), tp, fp, fn, tn, precision, recall, f1 } }
}
export type MaterialData = Awaited<ReturnType<typeof getMaterialData>>
