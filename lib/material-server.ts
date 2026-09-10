import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { parseWorkbookBuffer } from './material-data'

export async function getMaterialData() {
  const file = await readFile(path.join(process.cwd(), 'data/SIH26099_50row_dataset_with_ground_truth_labels-59b7df.xlsx'))
  return parseWorkbookBuffer(file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength), 'Demo material workbook')
}
