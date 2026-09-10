import { MaterialDashboard } from '@/components/material-dashboard'
import { getMaterialData } from '@/lib/material-data'

export default async function Page() {
  const data = await getMaterialData()
  return <MaterialDashboard data={data} />
}
