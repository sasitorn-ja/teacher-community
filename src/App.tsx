import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import * as htmlToImage from 'html-to-image'
import * as XLSX from 'xlsx'
import { ArrowRight, BookOpen, Camera, Check, ChevronDown, Database, Download, FileSpreadsheet, Image as ImageIcon, ImagePlus, KeyRound, LayoutDashboard, LoaderCircle, LockKeyhole, LogIn, LogOut, MapPin, Menu, Pencil, Plus, Search, Sparkles, Trash2, UserCog, Users, X } from 'lucide-react'
import { supabase, hasSupabaseConfig, supabaseConfigError } from './lib/supabase'
import { seededTeachers } from './data/teachers'
import './App.css'
import './account.css'
import './safety-theme.css'
import { Pagination, PaginationContent, PaginationEllipsis, PaginationItem, PaginationLink, PaginationNext, PaginationPrevious } from './components/ui/pagination'

type Community = { id:string; community_code:string; activity_date:string; community_name:string; advisor_name:string; school_name:string; location:string; member_count:number; description:string; image_url?:string; owner_id?:string }
type TeacherAccount = { id:string; teacher_name:string; community_code:string; role:'teacher'|'admin'; community_id?:string; login_email?:string }
type ImportedTeacherRow = { teacher_name:string; community_code:string }
const emptyForm = { activity_date:'', community_code:'', community_name:'', advisor_name:'', school_name:'', location:'', member_count:'', description:'', image_url:'' }
type CommunityForm = typeof emptyForm
type AdminCommunitySave = (owner:TeacherAccount,form:CommunityForm,imageFile:File|null,imagePreview:string)=>Promise<void>
const TEMPLATE_DETAIL_MAX_LENGTH = 120
const limitTemplateDetail = (value:string) => Array.from(value).slice(0,TEMPLATE_DETAIL_MAX_LENGTH).join('')
// Longer details still fit, but the text shrinks noticeably in the smaller template boxes.
const TEMPLATE_DETAIL_RECOMMENDED_LENGTH = 100
// Every field except the activity description is required, including the teacher photo shown on the template.
function missingCommunityFields(form:CommunityForm,hasImage:boolean) {
  const required:[string,string][] = [['ชื่อชุมนุม',form.community_name],['ครูที่ปรึกษา',form.advisor_name],['โรงเรียน',form.school_name],['สถานที่',form.location],['จำนวนที่รับ',form.member_count]]
  const missing = required.filter(([,value])=>!value.trim()).map(([label])=>label)
  if (!hasImage) missing.push('รูปภาพคุณครู')
  return missing
}
const missingFieldsMessage = (missing:string[]) => `กรุณากรอกให้ครบ: ${missing.join(', ')}`
type TemplateVariant = 'html'|'classic'|'colorful'|'modern'
type TemplateFieldKey = 'code'|'name'|'advisor'|'location'|'members'|'description'
// Image templates are laid out on the background's native 1254px grid, then scaled to fit their container.
const POSTER_SIZE = 1254
type PosterBox = { left:number; top:number; width:number; height:number }
type PosterField = { box:PosterBox; max:number; min:number; className?:string }
type TemplateVariantConfig = { label:string; background?:string; photo:PosterBox; fields:Record<TemplateFieldKey,PosterField> }
const posterDescription = 'poster-field-description'
const posterShort = 'poster-field-short'
// Cards drawn by the HTML template; each value sits inside its card, below the label chip.
const htmlPosterCards:Record<TemplateFieldKey,{label:string;tone:string;card:PosterBox}> = {
  code:{label:'รหัสชุมนุม',tone:'yellow',card:{left:572,top:300,width:622,height:104}},
  name:{label:'ชื่อชุมนุม',tone:'pink',card:{left:572,top:420,width:622,height:136}},
  advisor:{label:'ครูที่ปรึกษา',tone:'blue',card:{left:572,top:572,width:622,height:112}},
  location:{label:'สถานที่',tone:'green',card:{left:572,top:700,width:622,height:112}},
  members:{label:'จำนวนที่รับ',tone:'purple',card:{left:572,top:828,width:622,height:104}},
  description:{label:'รายละเอียด',tone:'orange',card:{left:572,top:948,width:622,height:196}},
}
const htmlPosterValue = ({left,top,width,height}:PosterBox):PosterBox => ({left:left+24,top:top+28,width:width-48,height:height-36})
const templateVariants:Record<TemplateVariant,TemplateVariantConfig> = {
  html:{label:'การ์ดพาสเทล',photo:{left:60,top:300,width:470,height:640},fields:{
    code:{box:htmlPosterValue(htmlPosterCards.code.card),max:44,min:22,className:posterShort},
    name:{box:htmlPosterValue(htmlPosterCards.name.card),max:40,min:18},
    advisor:{box:htmlPosterValue(htmlPosterCards.advisor.card),max:38,min:18},
    location:{box:htmlPosterValue(htmlPosterCards.location.card),max:34,min:15},
    members:{box:htmlPosterValue(htmlPosterCards.members.card),max:42,min:20,className:posterShort},
    description:{box:htmlPosterValue(htmlPosterCards.description.card),max:28,min:13,className:posterDescription},
  }},
  classic:{label:'ห้องเรียนน่ารัก',background:'/club-template-original.png',photo:{left:74,top:584,width:452,height:400},fields:{
    code:{box:{left:780,top:192,width:350,height:76},max:44,min:20,className:posterShort},
    name:{box:{left:650,top:362,width:540,height:88},max:40,min:18},
    advisor:{box:{left:650,top:545,width:540,height:95},max:38,min:18},
    location:{box:{left:665,top:740,width:525,height:58},max:32,min:15},
    members:{box:{left:665,top:882,width:525,height:60},max:40,min:18,className:posterShort},
    description:{box:{left:630,top:1040,width:560,height:80},max:24,min:12,className:posterDescription},
  }},
  colorful:{label:'ดอกไม้ท้องฟ้าใส',background:'/template-colorful.png',photo:{left:72,top:340,width:476,height:642},fields:{
    code:{box:{left:812,top:356,width:352,height:74},max:54,min:22,className:posterShort},
    name:{box:{left:604,top:500,width:580,height:76},max:40,min:18},
    advisor:{box:{left:604,top:650,width:556,height:64},max:44,min:18},
    location:{box:{left:604,top:792,width:566,height:50},max:32,min:15},
    members:{box:{left:822,top:878,width:332,height:68},max:50,min:20,className:posterShort},
    description:{box:{left:606,top:1032,width:576,height:96},max:26,min:13,className:posterDescription},
  }},
  modern:{label:'มินิมอล ฟ้า-ส้ม',background:'/template-modern.png',photo:{left:44,top:318,width:450,height:602},fields:{
    code:{box:{left:740,top:318,width:460,height:82},max:46,min:22,className:posterShort},
    name:{box:{left:735,top:438,width:465,height:94},max:40,min:18},
    advisor:{box:{left:745,top:572,width:455,height:102},max:38,min:18},
    location:{box:{left:720,top:710,width:480,height:82},max:34,min:15},
    members:{box:{left:755,top:828,width:445,height:84},max:42,min:20,className:posterShort},
    description:{box:{left:545,top:992,width:655,height:138},max:28,min:13,className:posterDescription},
  }},
}
const emptyAccount = { teacher_name:'', community_code:'' }
const currentAccountKey = 'teacher-community-current-account'
const adminTabStorageKey = 'teacher-community-admin-tab'
const todayInBangkok = () => new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Bangkok',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date())
const demos: Community[] = [
  { id:'demo-1', activity_date:todayInBangkok(), community_code:'กก026', community_name:'ร้านค้าสวัสดิการโรงเรียน', advisor_name:'นางจุฑาวรรณ ลิทธิชัย', school_name:'โรงเรียนวิเชียรมาตุ', location:'ร้านค้าสวัสดิการฯ', member_count:22, description:'เรียนรู้เกี่ยวกับการจัดการร้านค้าขายปลีก ฝึกการทำงานเป็นทีมและความรับผิดชอบ', owner_id:'teacher-1' },
  { id:'demo-2', activity_date:todayInBangkok(), community_code:'กก027', community_name:'นักสร้างสื่อสร้างสรรค์', advisor_name:'นายธนกร วัฒนชัย', school_name:'โรงเรียนวิเชียรมาตุ', location:'ห้องคอมพิวเตอร์ 2', member_count:22, description:'ฝึกออกแบบสื่อการเรียนรู้และสื่อประชาสัมพันธ์อย่างสร้างสรรค์', owner_id:'teacher-2' },
]
const demoAccounts: TeacherAccount[] = [{ id:'admin-1', teacher_name:'ผู้ดูแลระบบ', community_code:'admin2569', role:'admin' }, ...seededTeachers]
type SystemSettings = { title:string; term:string; year:string; school:string; adminUsername:string; adminPassword?:string }
const defaultSettings:SystemSettings = { title:'ระบบลงทะเบียนชุมนุมคุณครู', term:'ภาคเรียนที่ 2', year:'ปีการศึกษา 2569', school:'โรงเรียนวิเชียรมาตุ', adminUsername:'admin' }
const number = (n:number) => new Intl.NumberFormat('th-TH').format(n)
const thaiDate = (value:string) => value ? new Intl.DateTimeFormat('th-TH',{timeZone:'Asia/Bangkok',day:'numeric',month:'long',year:'numeric'}).format(new Date(`${value}T00:00:00+07:00`)) : '-'
const formatBytes = (bytes:number) => bytes < 1024 ? `${bytes} B` : bytes < 1024 * 1024 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / (1024 * 1024)).toFixed(2)} MB`
const normalizeSearch = (value:string) => value.normalize('NFKC').trim().toLowerCase().replace(/[\s-]+/g,'')
const normalizeCommunityCode = (value:string) => value.normalize('NFKC').replace(/\s+/g,'').trim()
const normalizeImportHeader = (value:string) => normalizeSearch(value).replace(/[.:：_()/\\]/g,'')
function extractTeacherRows(workbook:XLSX.WorkBook):ImportedTeacherRow[] {
  const teacherHeaders = ['ครูที่ปรึกษา','ครูผู้ดูแล','teacher','teachername','advisor','advisorname']
  const codeHeaders = ['รหัสชุมนุม','รหัสกิจกรรม','รหัส','communitycode','code']
  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName],{header:1,defval:'',raw:false}) as unknown[][]
    for (let headerRow=0; headerRow<Math.min(rows.length,80); headerRow += 1) {
      const headers = rows[headerRow].map((cell)=>normalizeImportHeader(String(cell)))
      const teacherIndex = headers.findIndex((header)=>teacherHeaders.some((candidate)=>header===normalizeImportHeader(candidate) || header.includes(normalizeImportHeader(candidate))))
      const codeIndex = headers.findIndex((header)=>codeHeaders.some((candidate)=>header===normalizeImportHeader(candidate) || header.includes(normalizeImportHeader(candidate))))
      if (teacherIndex < 0 || codeIndex < 0) continue
      const imported:ImportedTeacherRow[] = []
      for (const row of rows.slice(headerRow+1)) {
        const teacher_name = String(row[teacherIndex] ?? '').replace(/\s+/g,' ').trim()
        const community_code = String(row[codeIndex] ?? '').replace(/\s+/g,'').trim()
        if (!teacher_name || !community_code || teacher_name.includes('ครูที่ปรึกษา') || community_code === 'รหัสชุมนุม') continue
        imported.push({teacher_name,community_code})
      }
      const unique = Array.from(new Map(imported.map((item)=>[`${normalizeSearch(item.teacher_name)}:${normalizeSearch(item.community_code)}`,item])).values())
      if (unique.length) return unique
    }
  }
  throw new Error('ไม่พบคอลัมน์ “ครูที่ปรึกษา” และ “รหัสชุมนุม” ในไฟล์')
}
const errorText = (error:unknown,fallback:string) => error instanceof Error ? error.message : fallback
function withRequestTimeout<T>(operation:PromiseLike<T>,label:string,timeoutMs=15000):Promise<T> {
  return new Promise<T>((resolve,reject)=>{
    const timer = window.setTimeout(()=>reject(new Error(`${label}ใช้เวลานานเกินไป กรุณาลองใหม่`)),timeoutMs)
    Promise.resolve(operation).then((value)=>{window.clearTimeout(timer); resolve(value)},(error)=>{window.clearTimeout(timer); reject(error)})
  })
}
const fileToDataUrl = (file:File) => new Promise<string>((resolve,reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(new Error('อ่านไฟล์รูปไม่สำเร็จ')); reader.readAsDataURL(file) })
const compressImage = (file:File,maxBytes=500*1024) => new Promise<File>((resolve,reject) => {
  const reader = new FileReader()
  reader.onerror = () => reject(new Error('อ่านไฟล์รูปไม่สำเร็จ'))
  reader.onload = () => {
    const image = new Image()
    image.onerror = () => reject(new Error('เปิดไฟล์รูปไม่สำเร็จ'))
    image.onload = () => {
      const maxDimension = 1600
      const scale = Math.min(1,maxDimension / Math.max(image.width,image.height))
      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1,Math.round(image.width * scale)); canvas.height = Math.max(1,Math.round(image.height * scale))
      canvas.getContext('2d')?.drawImage(image,0,0,canvas.width,canvas.height)
      let quality = 0.84
      const convert = () => canvas.toBlob((blob) => {
        if (!blob) { reject(new Error('บีบอัดรูปไม่สำเร็จ')); return }
        if (blob.size <= maxBytes || quality <= 0.4) { resolve(new File([blob],`${file.name.replace(/\.[^.]+$/,'')}.jpg`,{type:'image/jpeg'})); return }
        quality -= 0.08; convert()
      },'image/jpeg',quality)
      convert()
    }
    image.src = String(reader.result)
  }
  reader.readAsDataURL(file)
})
type StorageItem = Pick<Community,'id'|'community_code'|'community_name'|'image_url'>
type StorageStats = { imageCount:number; imageUrlBytes:number; recordBytes:number }
type DatabaseUsageRow = { object_name:string; object_type:string; row_count:number; table_bytes:number; index_bytes:number; total_bytes:number }
type ApiAuditItem = { name:string; status:'ok'|'slow'|'error'|'not-connected'; durationMs:number; detail:string }
type ImagePickerController = { preview:string; choose:(file:File|null)=>void; clear:()=>void }
let imagePickerController:ImagePickerController|null = null

function getStorageObject(url:string) {
  try {
    const imageUrl = new URL(url)
    const projectUrl = import.meta.env.VITE_SUPABASE_URL as string | undefined
    if (!projectUrl || imageUrl.origin !== new URL(projectUrl).origin) return null
    const path = imageUrl.pathname
    const match = path.match(/\/storage\/v1\/object\/(?:public|sign|authenticated)\/([^/]+)\/(.+)$/)
    return match ? {bucket:decodeURIComponent(match[1]),path:decodeURIComponent(match[2])} : null
  } catch { return null }
}

function App() {
  const [communities,setCommunities] = useState<Community[]>([])
  const [accounts,setAccounts] = useState<TeacherAccount[]>([])
  const [current,setCurrent] = useState<TeacherAccount | null>(null)
  const [view,setView] = useState<'teacher'|'admin'>('teacher')
  const [adminTab,setAdminTab] = useState<AdminTab>(() => {
    if (typeof window === 'undefined') return 'template'
    return (window.sessionStorage.getItem(adminTabStorageKey) as AdminTab | null) ?? 'template'
  })
  const [form,setForm] = useState(emptyForm)
  const [editingId,setEditingId] = useState<string|null>(null)
  const [selectedId,setSelectedId] = useState<string|null>(null)
  const [accountForm,setAccountForm] = useState(emptyAccount)
  const [editingAccount,setEditingAccount] = useState<string|null>(null)
  const [importingAccounts,setImportingAccounts] = useState(false)
  const [teacherSearch,setTeacherSearch] = useState('')
  const [adminUsername,setAdminUsername] = useState('')
  const [search,setSearch] = useState('')
  const [showLogin,setShowLogin] = useState(true)
  const [showAdminLogin,setShowAdminLogin] = useState(false)
  const [settings,setSettings] = useState<SystemSettings>(defaultSettings)
  const [notice,setNotice] = useState('')
  const [connectionError,setConnectionError] = useState('')
  const [loading,setLoading] = useState(true)
  const [saving,setSaving] = useState(false)
  const [todaySubmissionCount,setTodaySubmissionCount] = useState(0)
  const [today,setToday] = useState(todayInBangkok())
  const [imageFile,setImageFile] = useState<File|null>(null)
  const [imagePreview,setImagePreview] = useState('')
  const cardRef = useRef<HTMLDivElement>(null)
  const clearFormAfterSaveRef = useRef(false)
  const isAdmin = current?.role === 'admin'
  const selected = communities.find((item) => item.id === selectedId) ?? communities[0]
  const filtered = useMemo(() => communities.filter((item) => [item.community_code,item.community_name,item.advisor_name].join(' ').toLowerCase().includes(search.toLowerCase())),[communities,search])
  const storageItems = useMemo<StorageItem[]>(() => communities.filter((item) => Boolean(item.image_url?.trim())),[communities])
  const storageStats = useMemo<StorageStats>(() => ({
    imageCount: storageItems.length,
    imageUrlBytes: storageItems.reduce((total,item) => total + new TextEncoder().encode(item.image_url ?? '').length,0),
    recordBytes: communities.reduce((total,item) => total + new TextEncoder().encode([item.community_code,item.community_name,item.advisor_name,item.school_name,item.location,item.description,item.image_url ?? ''].join('|')).length,0),
  }),[communities,storageItems])
  const teacherMatches = useMemo(() => {
    const needle = normalizeSearch(teacherSearch)
    return accounts.filter((item) => item.role==='teacher' && (!needle || [item.teacher_name,item.community_code,item.login_email ?? ''].some((value)=>normalizeSearch(value).includes(needle)))).slice(0,9)
  },[accounts,teacherSearch])
  useEffect(() => { localStorage.removeItem('teacher-community-remembered-name'); void load() }, [])
  useEffect(() => { const timer = window.setInterval(() => setToday(todayInBangkok()),60000); return () => window.clearInterval(timer) }, [])
  useEffect(() => {
    if (!hasSupabaseConfig || !supabase) return
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event) => {
      if (event === 'SIGNED_OUT') {
        setCurrent(null); setTodaySubmissionCount(0); setShowLogin(true); setView('teacher')
      }
    })
    return () => subscription.unsubscribe()
  }, [])
  useEffect(() => { if (current && current.role !== 'admin') void refreshSubmissionCount(current.id,today) }, [current,today])
  useEffect(() => {
    if (!current || current.role === 'admin') return
    if (clearFormAfterSaveRef.current) {
      clearFormAfterSaveRef.current = false
      return
    }
    const record = communities.find((item) => item.owner_id === current.id && item.activity_date === today)
    if (record) {
      setEditingId(record.id)
      setForm({activity_date:record.activity_date,community_code:record.community_code,community_name:record.community_name,advisor_name:record.advisor_name,school_name:record.school_name,location:record.location,member_count:String(record.member_count),description:limitTemplateDetail(record.description),image_url:record.image_url ?? ''})
      setImageFile(null); setImagePreview(record.image_url ?? '')
    } else {
      setEditingId(null)
      setForm({...emptyForm,activity_date:today,community_code:current.community_code,advisor_name:current.teacher_name,school_name:settings.school})
      setImageFile(null); setImagePreview('')
    }
  },[current,communities,today,settings.school])
  const flash = (text:string) => { setNotice(text); window.setTimeout(() => setNotice(''),4500) }
  async function refreshSubmissionCount(ownerId:string,date:string) {
    if (!hasSupabaseConfig || !supabase) {
      const stored = JSON.parse(localStorage.getItem('teacher-community-submission-counts') ?? '{}') as Record<string,number>
      const count = Number(stored[`${ownerId}:${date}`] ?? 0)
      setTodaySubmissionCount(count)
      return count
    }
    try {
      const result = await withRequestTimeout(supabase.from('community_submission_events').select('id',{count:'exact',head:true}).eq('owner_id',ownerId).eq('activity_date',date),'โหลดจำนวนครั้งที่ส่ง')
      if (result.error) throw result.error
      const count = result.count ?? 0
      setTodaySubmissionCount(count)
      return count
    } catch {
      setTodaySubmissionCount(0)
      return 0
    }
  }
  async function loadDatabaseUsage() {
    if (!hasSupabaseConfig || !supabase) throw new Error('ยังไม่ได้เชื่อมต่อ Supabase')
    const result = await withRequestTimeout(supabase.rpc('get_database_usage'),'โหลดพื้นที่ฐานข้อมูล',15000)
    if (result.error) throw result.error
    return ((result.data ?? []) as DatabaseUsageRow[]).map((row)=>({...row,row_count:Number(row.row_count),table_bytes:Number(row.table_bytes),index_bytes:Number(row.index_bytes),total_bytes:Number(row.total_bytes)}))
  }
  async function runApiAudit():Promise<ApiAuditItem[]> {
    if (!hasSupabaseConfig || !supabase) return [{name:'Supabase',status:'not-connected',durationMs:0,detail:'ยังไม่ได้ตั้งค่า URL หรือ anon key'}]
    const check = async (name:string,operation:()=>PromiseLike<unknown>):Promise<ApiAuditItem> => {
      const started = performance.now()
      try {
        const result = await withRequestTimeout(operation(),`ตรวจสอบ ${name}`,10000) as {error?:{message?:string}|null}
        const durationMs = Math.round(performance.now()-started)
        if (result?.error) throw new Error(result.error.message ?? 'API ตอบกลับเป็นข้อผิดพลาด')
        return {name,status:durationMs>1000?'slow':'ok',durationMs,detail:durationMs>1000?'เชื่อมต่อได้ แต่ตอบสนองช้า':'เชื่อมต่อปกติ'}
      } catch (error) {
        return {name,status:'error',durationMs:Math.round(performance.now()-started),detail:errorText(error,'เชื่อมต่อไม่สำเร็จ')}
      }
    }
    return Promise.all([
      check('ฐานข้อมูล: communities',()=>supabase!.from('communities').select('id',{count:'exact',head:true})),
      check('ฐานข้อมูล: teacher_profiles RPC',()=>supabase!.rpc('list_teacher_directory')),
      check('ฐานข้อมูล: teacher_profiles',()=>supabase!.from('teacher_profiles').select('id',{count:'exact',head:true})),
      check('ประวัติการส่งข้อมูล',()=>supabase!.from('community_submission_events').select('id',{count:'exact',head:true})),
      check('ฐานข้อมูล: system_settings',()=>supabase!.from('system_settings').select('id').eq('id','main').maybeSingle()),
      check('Storage: community-images',()=>supabase!.storage.from('community-images').list('',{limit:1})),
      check('รายงานพื้นที่ฐานข้อมูล',()=>supabase!.rpc('get_database_usage')),
      check('Edge Function: admin-accounts',async()=>{
        const response = await fetch(`${String(import.meta.env.VITE_SUPABASE_URL).replace(/\/$/,'')}/functions/v1/admin-accounts`,{method:'OPTIONS'})
        if (!response.ok) throw new Error(`ตอบกลับ HTTP ${response.status}`)
        return response
      }),
    ])
  }
  async function recordSubmission(ownerId:string,communityId:string,date:string) {
    if (!hasSupabaseConfig || !supabase) {
      const stored = JSON.parse(localStorage.getItem('teacher-community-submission-counts') ?? '{}') as Record<string,number>
      const key = `${ownerId}:${date}`
      stored[key] = Number(stored[key] ?? 0) + 1
      localStorage.setItem('teacher-community-submission-counts',JSON.stringify(stored))
      return refreshSubmissionCount(ownerId,date)
    }
    const result = await withRequestTimeout(supabase.from('community_submission_events').insert({owner_id:ownerId,community_id:communityId,activity_date:date}),'บันทึกประวัติการส่ง')
    if (result.error) throw result.error
    return refreshSubmissionCount(ownerId,date)
  }
  async function load(showBootstrap=true) {
    if (showBootstrap) setLoading(true)
    setConnectionError('')
    if (!hasSupabaseConfig && import.meta.env.PROD) {
      setCommunities([]); setAccounts([]); setCurrent(null); setTodaySubmissionCount(0); setShowLogin(true); setConnectionError(supabaseConfigError); setLoading(false); return
    }
    if (!hasSupabaseConfig) {
      const storedCommunities = JSON.parse(localStorage.getItem('teacher-community-demo') ?? 'null') as Community[] | null
      const localCommunities = storedCommunities?.map((item) => ({...item,activity_date:item.activity_date || todayInBangkok(),school_name:defaultSettings.school})) ?? null
      const localAccounts = JSON.parse(localStorage.getItem('teacher-community-accounts') ?? 'null') as TeacherAccount[] | null
      const migratedAccounts = !localAccounts || localAccounts.filter((item)=>item.role==='teacher').length < 100 ? demoAccounts : localAccounts
      const savedSettings = JSON.parse(localStorage.getItem('teacher-community-settings') ?? 'null') as SystemSettings | null
      const localSettings = savedSettings ? {...defaultSettings,...savedSettings,school:defaultSettings.school,adminUsername:savedSettings.adminUsername || defaultSettings.adminUsername} : null
      const restored = migratedAccounts.find((item) => item.id === localStorage.getItem(currentAccountKey)) ?? null
      const localData = localCommunities ?? demos
      setCommunities(localData); setAccounts(migratedAccounts); setSettings(localSettings ?? defaultSettings); setSelectedId(localData[0]?.id ?? null); setCurrent(restored); setView(restored?.role === 'admin' ? 'admin' : 'teacher'); setShowLogin(!restored); if (restored?.role === 'teacher') void refreshSubmissionCount(restored.id,today); setLoading(false); return
    }
    try {
      const [{data: dataCommunities,error:communityError},{data:{session},error:sessionError},directoryResult,settingsResult] = await Promise.all([
        withRequestTimeout(supabase!.from('communities').select('*').order('created_at'),'โหลดข้อมูลชุมชน'),
        withRequestTimeout(supabase!.auth.getSession(),'กู้คืนเซสชันการเข้าสู่ระบบ'),
        withRequestTimeout(supabase!.rpc('list_teacher_directory'),'โหลดรายชื่อครู'),
        withRequestTimeout(supabase!.from('system_settings').select('id,title,term,year,school').eq('id','main').maybeSingle(),'โหลดค่าระบบ'),
      ])
      if (communityError) throw new Error(`โหลดข้อมูลชุมชนไม่สำเร็จ: ${communityError.message}`)
      if (sessionError) throw sessionError
      if (directoryResult.error) throw new Error(`โหลดรายชื่อครูไม่สำเร็จ: ${directoryResult.error.message}`)
      if (settingsResult.error) throw new Error(`โหลดค่าระบบไม่สำเร็จ: ${settingsResult.error.message}`)
      const fixedCommunities = (dataCommunities ?? []) as Community[]
      setCommunities(fixedCommunities); setSelectedId(fixedCommunities[0]?.id ?? null)
      const directoryAccounts = ((directoryResult.data ?? []) as Array<{id:string;teacher_name:string}>).map((item)=>({id:item.id,teacher_name:item.teacher_name,community_code:'',role:'teacher' as const}))
      if (directoryAccounts.length) setAccounts(directoryAccounts)
      const loadedSettings = settingsResult.data as Partial<SystemSettings> | null
      if (loadedSettings) setSettings({...defaultSettings,...loadedSettings,adminUsername:defaultSettings.adminUsername})
      if (session?.user) {
        const {data: account,error:profileError} = await withRequestTimeout(supabase!.from('teacher_profiles').select('*').eq('id',session.user.id).single(),'โหลดโปรไฟล์ผู้ใช้')
        if (profileError) throw new Error(`โหลดโปรไฟล์ผู้ใช้ไม่สำเร็จ: ${profileError.message}`)
        const restored = account as TeacherAccount | null
        setCurrent(restored); setShowLogin(!restored); setView(restored?.role === 'admin' ? 'admin' : 'teacher')
        if (restored?.role === 'teacher') void refreshSubmissionCount(restored.id,today)
        if (restored?.role === 'admin') setSettings((currentSettings)=>({...currentSettings,adminUsername:restored.teacher_name}))
        if (account?.role === 'admin') {
          const {data,error} = await withRequestTimeout(supabase!.from('teacher_profiles').select('*').order('teacher_name'),'โหลดบัญชีครู')
          if (error) throw new Error(`โหลดบัญชีครูไม่สำเร็จ: ${error.message}`)
          setAccounts((data ?? []) as TeacherAccount[])
        }
      } else { setCurrent(null); setTodaySubmissionCount(0); setShowLogin(true); setView('teacher') }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'เชื่อมต่อ Supabase ไม่สำเร็จ'
      setConnectionError(message); setCommunities([]); setAccounts([]); setCurrent(null); setTodaySubmissionCount(0); setShowLogin(true); setView('teacher')
    } finally {
      setLoading(false)
    }
  }
  function setField(key:keyof typeof emptyForm,value:string) { setForm((now) => ({...now,[key]:key==='description'?limitTemplateDetail(value):key==='school_name'?settings.school:key==='advisor_name'?(current?.teacher_name ?? value):value})) }
  async function chooseImage(file:File|null) {
    if (!file) return
    if (!file.type.startsWith('image/')) { flash('กรุณาเลือกไฟล์รูปภาพเท่านั้น'); return }
    if (file.size > 12 * 1024 * 1024) { flash('รูปภาพต้นฉบับต้องมีขนาดไม่เกิน 12 MB'); return }
    let compactFile:File
    try { compactFile = await compressImage(file) } catch (error) { flash(error instanceof Error ? error.message : 'บีบอัดรูปไม่สำเร็จ'); return }
    if (imagePreview.startsWith('blob:')) URL.revokeObjectURL(imagePreview)
    setImageFile(compactFile); setImagePreview(URL.createObjectURL(compactFile)); flash(`ย่อรูปเหลือ ${formatBytes(compactFile.size)} แล้ว`)
  }
  function clearImage() {
    if (imagePreview.startsWith('blob:')) URL.revokeObjectURL(imagePreview)
    setImageFile(null); setImagePreview(''); setField('image_url','')
  }
  async function getImageUrlForSave() {
    if (!imageFile) return imagePreview || form.image_url.trim() || null
    if (!hasSupabaseConfig) return fileToDataUrl(imageFile)
    const extension = imageFile.name.split('.').pop()?.toLowerCase().replace(/[^a-z0-9]/g,'') || 'jpg'
    const path = `${current!.id}/${today}-${crypto.randomUUID()}.${extension}`
    const upload = await withRequestTimeout(supabase!.storage.from('community-images').upload(path,imageFile,{contentType:imageFile.type,upsert:false}),'อัปโหลดรูป')
    if (upload.error) throw new Error(`อัปโหลดรูปไม่สำเร็จ: ${upload.error.message}`)
    return supabase!.storage.from('community-images').getPublicUrl(path).data.publicUrl
  }
  function openEditor(community?:Community) {
    if (!current && hasSupabaseConfig) { setShowLogin(true); flash('กรุณาเข้าสู่ระบบก่อนแก้ไขข้อมูล'); return }
    if (community) { setEditingId(community.id); setImageFile(null); setImagePreview(community.image_url ?? ''); setForm({activity_date:community.activity_date,community_code:community.community_code,community_name:community.community_name,advisor_name:community.advisor_name,school_name:community.school_name,location:community.location,member_count:String(community.member_count),description:limitTemplateDetail(community.description),image_url:community.image_url ?? ''}) } else { setEditingId(null); setImageFile(null); setImagePreview(''); setForm({...emptyForm,activity_date:today,community_code:current?.community_code ?? '',advisor_name:current?.teacher_name ?? '',school_name:settings.school}) }
    setView('teacher'); window.scrollTo({top:0,behavior:'smooth'})
  }
  function clearSubmittedCommunityForm() {
    if (imagePreview.startsWith('blob:')) URL.revokeObjectURL(imagePreview)
    setEditingId(null)
    setImageFile(null)
    setImagePreview('')
    setForm({...emptyForm,activity_date:today,community_code:current?.community_code ?? '',advisor_name:current?.teacher_name ?? '',school_name:settings.school})
  }
  async function saveCommunity(event:React.FormEvent) {
    event.preventDefault()
    if (!current) { setShowLogin(true); flash('กรุณาเข้าสู่ระบบก่อนบันทึก'); return }
    const missing = missingCommunityFields(form,Boolean(imageFile || imagePreview || form.image_url.trim()))
    if (missing.length) { flash(missingFieldsMessage(missing)); return }
    setSaving(true)
    try {
      let imageUrl:string|null
      try { imageUrl = await getImageUrlForSave() } catch (error) { flash(error instanceof Error ? error.message : 'อัปโหลดรูปไม่สำเร็จ'); return }
      const memberCount = Number(form.member_count)
      if (!Number.isInteger(memberCount) || memberCount < 22) { flash('จำนวนที่รับต้องอย่างน้อย 22 คน และไม่จำกัดจำนวนสูงสุด'); return }
      const payload = {activity_date:today,community_name:form.community_name.trim(),advisor_name:current.teacher_name,school_name:form.school_name.trim(),location:form.location.trim(),member_count:memberCount,description:limitTemplateDetail(form.description.trim()),image_url:imageUrl}
      if (!hasSupabaseConfig) {
        const code = current.community_code
        const existing = communities.find((x)=>x.owner_id===current.id && x.activity_date===today)
        const recordId = editingId ?? existing?.id
        const item:Community = recordId ? {...communities.find((x)=>x.id===recordId)!, ...payload, community_code:code, image_url:payload.image_url ?? undefined} : {...payload,id:crypto.randomUUID(),community_code:code,owner_id:current.id,image_url:payload.image_url ?? undefined}
        const next = recordId ? communities.map((x)=>x.id===item.id?item:x) : [...communities,item]
        clearFormAfterSaveRef.current = true
        setCommunities(next); localStorage.setItem('teacher-community-demo',JSON.stringify(next));
        if (!recordId) { const nextAccounts=accounts.map((x)=>x.id===current.id?{...x,community_id:item.id}:x); setAccounts(nextAccounts); setCurrent(nextAccounts.find((x)=>x.id===current.id) ?? current); localStorage.setItem('teacher-community-accounts',JSON.stringify(nextAccounts)) }
        setSelectedId(item.id); clearSubmittedCommunityForm()
        const count = await recordSubmission(current.id,item.id,today)
        flash(`บันทึกสำเร็จแล้ว · วันนี้ส่งข้อมูลแล้ว ${number(count)} ครั้ง`)
        return
      }
      const existing = communities.find((x)=>x.owner_id===current.id && x.activity_date===today)
      const recordId = editingId ?? existing?.id
      let answer
      try { answer = recordId ? await withRequestTimeout(supabase!.from('communities').update(payload).eq('id',recordId).select().single(),'บันทึกข้อมูล') : await withRequestTimeout(supabase!.from('communities').insert({...payload,community_code:current.community_code,owner_id:current.id}).select().single(),'บันทึกข้อมูล') }
      catch (error) { flash(`บันทึกไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
      if (answer.error) { flash(`บันทึกไม่สำเร็จ: ${answer.error.message}`); return }
      const count = await recordSubmission(current.id,answer.data.id,today)
      clearFormAfterSaveRef.current = true
      await load(false); setSelectedId(answer.data.id); clearSubmittedCommunityForm(); flash(`บันทึกสำเร็จแล้ว · วันนี้ส่งข้อมูลแล้ว ${number(count)} ครั้ง`)
    } finally {
      setSaving(false)
    }
  }
  async function saveCommunityForAdmin(owner:TeacherAccount,adminForm:CommunityForm,adminImageFile:File|null,adminImagePreview:string) {
    const missing = missingCommunityFields(adminForm,Boolean(adminImageFile || adminImagePreview || adminForm.image_url.trim()))
    if (missing.length) { flash(missingFieldsMessage(missing)); return }
    const memberCount = Number(adminForm.member_count)
    if (!Number.isInteger(memberCount) || memberCount < 22) { flash('จำนวนที่รับต้องอย่างน้อย 22 คน และไม่จำกัดจำนวนสูงสุด'); return }
    let imageUrl:string|null = adminImagePreview || adminForm.image_url.trim() || null
    try {
      if (adminImageFile) {
        if (!hasSupabaseConfig) imageUrl = await fileToDataUrl(adminImageFile)
        else {
          const extension = adminImageFile.name.split('.').pop()?.toLowerCase().replace(/[^a-z0-9]/g,'') || 'jpg'
          const path = `${owner.id}/${adminForm.activity_date || today}-${crypto.randomUUID()}.${extension}`
          const upload = await withRequestTimeout(supabase!.storage.from('community-images').upload(path,adminImageFile,{contentType:adminImageFile.type,upsert:false}),'อัปโหลดรูป')
          if (upload.error) throw new Error(`อัปโหลดรูปไม่สำเร็จ: ${upload.error.message}`)
          imageUrl = supabase!.storage.from('community-images').getPublicUrl(path).data.publicUrl
        }
      }
    } catch (error) { flash(error instanceof Error ? error.message : 'อัปโหลดรูปไม่สำเร็จ'); return }
    const activityDate = adminForm.activity_date || today
    const payload = {activity_date:activityDate,community_name:adminForm.community_name.trim(),advisor_name:owner.teacher_name,school_name:adminForm.school_name.trim(),location:adminForm.location.trim(),member_count:memberCount,description:limitTemplateDetail(adminForm.description.trim()),image_url:imageUrl}
    if (!hasSupabaseConfig) {
      const existing = communities.find((item)=>item.owner_id===owner.id && item.activity_date===activityDate)
      const item:Community = existing ? {...existing,...payload,community_code:owner.community_code,owner_id:owner.id,image_url:imageUrl ?? undefined} : {...payload,id:crypto.randomUUID(),community_code:owner.community_code,owner_id:owner.id,image_url:imageUrl ?? undefined}
      const next = existing ? communities.map((value)=>value.id===item.id?item:value) : [...communities,item]
      setCommunities(next); setSelectedId(item.id); localStorage.setItem('teacher-community-demo',JSON.stringify(next)); flash(existing?'แก้ไขข้อมูลชุมชนของคุณครูแล้ว':'บันทึกข้อมูลชุมชนแทนคุณครูแล้ว'); return
    }
    const existing = communities.find((item)=>item.owner_id===owner.id && item.activity_date===activityDate)
    let answer
    try { answer = existing ? await withRequestTimeout(supabase!.from('communities').update(payload).eq('id',existing.id).select().single(),'บันทึกข้อมูล') : await withRequestTimeout(supabase!.from('communities').insert({...payload,community_code:owner.community_code,owner_id:owner.id}).select().single(),'บันทึกข้อมูล') }
    catch (error) { flash(`บันทึกไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    if (answer.error) { flash(`บันทึกไม่สำเร็จ: ${answer.error.message}`); return }
    await load(false); flash(existing?'บันทึกสำเร็จแล้ว · แก้ไขข้อมูลชุมชนของคุณครูแล้ว':'บันทึกสำเร็จแล้ว · บันทึกข้อมูลชุมชนแทนคุณครูแล้ว')
  }
  async function signInWithSupabase(username:string,adminOnly=false):Promise<boolean> {
    const result = await withRequestTimeout(supabase!.functions.invoke('passwordless-login',{body:{username,adminOnly}}),'เข้าสู่ระบบ')
    if (result.error) { flash(`เข้าสู่ระบบไม่สำเร็จ: ${result.error.message}`); return false }
    const tokenHash = String((result.data as {token_hash?:string} | null)?.token_hash ?? '').trim()
    if (!tokenHash) { flash('เข้าสู่ระบบไม่สำเร็จ: ไม่ได้รับ token สำหรับสร้าง session'); return false }
    const {data,error} = await withRequestTimeout(supabase!.auth.verifyOtp({token_hash:tokenHash,type:'magiclink'}),'สร้าง session ผู้ใช้')
    if (error || !data.user) { flash(`เข้าสู่ระบบไม่สำเร็จ: ${error?.message ?? 'สร้าง session ไม่สำเร็จ'}`); return false }
    const profileResult = await withRequestTimeout(supabase!.from('teacher_profiles').select('*').eq('id',data.user.id).single(),'โหลดโปรไฟล์ผู้ใช้')
    if (profileResult.error || !profileResult.data) { flash(`โหลดโปรไฟล์ไม่สำเร็จ: ${profileResult.error?.message ?? 'ไม่พบโปรไฟล์'}`); await supabase!.auth.signOut(); return false }
    const account = profileResult.data as TeacherAccount
    if (adminOnly && account.role !== 'admin') { flash('บัญชีนี้ไม่มีสิทธิ์ผู้ดูแล'); await supabase!.auth.signOut(); return false }
    setCurrent(account); setShowLogin(false); setShowAdminLogin(false); setView(account.role === 'admin' ? 'admin' : 'teacher')
    localStorage.setItem(currentAccountKey,account.id)
    if (account.role === 'admin') {
      const {data:allAccounts,error:accountsError} = await withRequestTimeout(supabase!.from('teacher_profiles').select('*').order('teacher_name'),'โหลดบัญชีครู')
      if (accountsError) { flash(`โหลดบัญชีครูไม่สำเร็จ: ${accountsError.message}`); return false }
      setAccounts((allAccounts ?? []) as TeacherAccount[])
    }
    flash(`ยินดีต้อนรับ ${account.teacher_name}`)
    return true
  }

  function enterLocalAdmin() {
    const admin = accounts.find((item)=>item.role==='admin') ?? {id:'admin-1',teacher_name:'ผู้ดูแลระบบ',community_code:'',role:'admin' as const}
    localStorage.setItem(currentAccountKey,admin.id)
    setCurrent(admin); setShowLogin(false); setShowAdminLogin(false); setTeacherSearch(''); setAdminUsername(''); setView('admin')
    flash('เข้าสู่ระบบผู้ดูแลแล้ว · โหมด local สำหรับ localhost')
  }

  async function teacherSignIn(event:React.FormEvent) {
    event.preventDefault()
    const username = teacherSearch.trim()
    if (!username) { flash('กรุณากรอก Username'); return }
    if (hasSupabaseConfig) {
      try {
        const signedIn = await signInWithSupabase(username)
        if (!signedIn && import.meta.env.DEV && username.toLowerCase() === 'admin') enterLocalAdmin()
      } catch (error) {
        if (import.meta.env.DEV && username.toLowerCase() === 'admin') enterLocalAdmin()
        else flash(`เข้าสู่ระบบไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`)
      }
      return
    }
    if (username.toLowerCase() === settings.adminUsername.trim().toLowerCase()) {
      enterLocalAdmin()
      return
    }
    const account = accounts.find((item) => item.role === 'teacher' && normalizeSearch(item.teacher_name) === normalizeSearch(username))
    if (!account) { flash('ไม่พบชื่อผู้ใช้งาน กรุณาพิมพ์ชื่อแล้วเลือกจากรายชื่อ'); return }
    selectTeacher(account)
  }
  function selectTeacher(account:TeacherAccount) {
    if (hasSupabaseConfig) { setTeacherSearch(account.teacher_name); return }
    localStorage.setItem(currentAccountKey,account.id); setCurrent(account); setShowLogin(false); setTeacherSearch(''); flash(`ยินดีต้อนรับ ${account.teacher_name}`)
  }
  async function adminSignIn(event:React.FormEvent) {
    event.preventDefault()
    if (hasSupabaseConfig) {
      try {
        const signedIn = await signInWithSupabase(adminUsername.trim(),true)
        if (!signedIn && import.meta.env.DEV && adminUsername.trim().toLowerCase() === 'admin') enterLocalAdmin()
      } catch (error) {
        if (import.meta.env.DEV && adminUsername.trim().toLowerCase() === 'admin') enterLocalAdmin()
        else flash(`เข้าสู่ระบบไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`)
      }
      return
    }
    if (adminUsername.trim().toLowerCase() !== settings.adminUsername.trim().toLowerCase()) { flash('ไม่พบ Username ผู้ดูแลระบบ'); return }
    enterLocalAdmin()
  }
  async function updateSettings(next:SystemSettings) {
    const fixedSettings = {...next,school:next.school.trim() || defaultSettings.school}
    if (!hasSupabaseConfig) { setSettings(fixedSettings); localStorage.setItem('teacher-community-settings',JSON.stringify(fixedSettings)); flash('บันทึกชื่อระบบแล้ว'); return }
    let result
    try { result = await withRequestTimeout(supabase!.from('system_settings').upsert({id:'main',title:fixedSettings.title.trim(),term:fixedSettings.term.trim(),year:fixedSettings.year.trim(),school:fixedSettings.school}).select().single(),'บันทึกค่าระบบ') }
    catch (error) { flash(`บันทึกค่าระบบไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    const {error} = result
    if (error) { flash(`บันทึกค่าระบบไม่สำเร็จ: ${error.message}`); return }
    setSettings(fixedSettings)
    flash('บันทึกชื่อระบบแล้ว')
  }
  async function updateAdminAuth(next:SystemSettings) {
    if (!hasSupabaseConfig) { await updateSettings(next); return }
    if (!current || current.role !== 'admin') { flash('กรุณาเข้าสู่ระบบผู้ดูแลก่อน'); return }
    if (!next.adminUsername.trim()) { flash('กรุณากรอก Username ผู้ดูแล'); return }
    let profileResult
    try { profileResult = await withRequestTimeout(supabase!.from('teacher_profiles').update({teacher_name:next.adminUsername.trim()}).eq('id',current.id).select().single(),'บันทึกชื่อผู้ดูแล') }
    catch (error) { flash(`บันทึกชื่อผู้ดูแลไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    if (profileResult.error) { flash(`บันทึกชื่อผู้ดูแลไม่สำเร็จ: ${profileResult.error.message}`); return }
    const updated = profileResult.data as TeacherAccount
    setCurrent(updated); setSettings((currentSettings)=>({...currentSettings,adminUsername:updated.teacher_name})); flash('บันทึก Username ผู้ดูแลใน Supabase แล้ว')
  }
  // Keep this handler available to the admin-account tab in both local and Supabase modes.
  void updateAdminAuth
  async function deleteCommunityImage(community:Community) {
    if (!community.image_url) return
    if (!window.confirm(`ลบรูปของชุมนุม “${community.community_name}” ใช่หรือไม่?`)) return
    if (!hasSupabaseConfig) {
      const next = communities.map((item) => item.id === community.id ? {...item,image_url:undefined} : item)
      setCommunities(next); localStorage.setItem('teacher-community-demo',JSON.stringify(next)); flash('ลบรูปออกจากข้อมูลแล้ว'); return
    }
    const object = getStorageObject(community.image_url)
    let imageUpdate
    try { imageUpdate = await withRequestTimeout(supabase!.from('communities').update({image_url:null}).eq('id',community.id),'ลบลิงก์รูป') }
    catch (error) { flash(`ลบลิงก์รูปไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    const {error} = imageUpdate
    if (error) { flash(`ลบลิงก์รูปไม่สำเร็จ: ${error.message}`); return }
    await load(false); flash(object ? 'ลบข้อมูลอ้างอิงแล้ว กำลังเคลียร์ไฟล์รูปใน Storage' : 'ลบลิงก์รูปออกจากข้อมูลแล้ว')
    if (object) void withRequestTimeout(supabase!.storage.from(object.bucket).remove([object.path]),'ลบไฟล์รูป').then((answer)=>{
      if (answer.error && !/not found|does not exist/i.test(answer.error.message)) flash(`ลบไฟล์รูปไม่สำเร็จ: ${answer.error.message}`)
    }).catch((error)=>flash(error instanceof Error ? error.message : 'ลบไฟล์รูปใช้เวลานานเกินไป'))
  }
  async function deleteCommunitiesByDate(activityDate:string) {
    const matches = communities.filter((item)=>item.activity_date===activityDate)
    if (!matches.length) { flash('ไม่พบข้อมูลของวันที่เลือก'); return }
    if (!window.confirm(`ยืนยันลบข้อมูลชุมชนทั้งหมด ${matches.length} รายการ ของวันที่ ${thaiDate(activityDate)} ใช่หรือไม่?\n\nการลบนี้ไม่สามารถกู้คืนได้`)) return
    if (!hasSupabaseConfig) {
      const next = communities.filter((item)=>item.activity_date!==activityDate)
      setCommunities(next); setSelectedId(next[0]?.id ?? null); localStorage.setItem('teacher-community-demo',JSON.stringify(next)); flash(`ลบข้อมูลของวันที่ ${thaiDate(activityDate)} แล้ว`); return
    }
    const storedImages = matches.map((item)=>getStorageObject(item.image_url ?? '')).filter((item):item is {bucket:string;path:string}=>Boolean(item))
    let answer
    try { answer = await withRequestTimeout(supabase!.from('communities').delete().eq('activity_date',activityDate),'ลบข้อมูลชุมชน') }
    catch (error) { flash(`ลบข้อมูลไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    if (answer.error) { flash(`ลบข้อมูลไม่สำเร็จ: ${answer.error.message}`); return }
    await load(false); flash(`ลบข้อมูลของวันที่ ${thaiDate(activityDate)} จากฐานข้อมูลแล้ว กำลังเคลียร์ไฟล์รูป`)
    void Promise.allSettled(storedImages.map((item)=>withRequestTimeout(supabase!.storage.from(item.bucket).remove([item.path]),'ลบไฟล์รูป')))
  }
  async function deleteCommunity(community:Community) {
    if (!window.confirm(`ยืนยันลบข้อมูลชุมชน “${community.community_name}” ของ ${community.advisor_name} ใช่หรือไม่?\n\nการลบนี้ไม่สามารถกู้คืนได้`)) return
    if (!hasSupabaseConfig) {
      const next = communities.filter((item)=>item.id!==community.id)
      setCommunities(next); setSelectedId((selectedId===community.id ? next[0]?.id : selectedId) ?? null); localStorage.setItem('teacher-community-demo',JSON.stringify(next)); flash(`ลบข้อมูลชุมชน “${community.community_name}” แล้ว`); return
    }
    const storedImage = getStorageObject(community.image_url ?? '')
    let answer
    try { answer = await withRequestTimeout(supabase!.from('communities').delete().eq('id',community.id),'ลบข้อมูลชุมชน') }
    catch (error) { flash(`ลบข้อมูลไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    if (answer.error) { flash(`ลบข้อมูลไม่สำเร็จ: ${answer.error.message}`); return }
    await load(false); flash(`ลบข้อมูลชุมชน “${community.community_name}” จากฐานข้อมูลแล้ว`)
    if (storedImage) void withRequestTimeout(supabase!.storage.from(storedImage.bucket).remove([storedImage.path]),'ลบไฟล์รูป').then((result)=>{ if (result.error && !/not found|does not exist/i.test(result.error.message)) flash(`ลบไฟล์รูปไม่สำเร็จ: ${result.error.message}`) }).catch((error)=>flash(error instanceof Error ? error.message : 'ลบไฟล์รูปใช้เวลานานเกินไป'))
  }
  async function deleteCommunitiesBeforeDate(cutoffDate:string) {
    const matches = communities.filter((item)=>item.activity_date && item.activity_date < cutoffDate)
    if (!cutoffDate || !matches.length) { flash('ไม่พบข้อมูลเก่าก่อนวันที่เลือก'); return }
    if (!window.confirm(`ยืนยันลบข้อมูลเก่าก่อนวันที่ ${thaiDate(cutoffDate)} จำนวน ${matches.length} รายการ ใช่หรือไม่?\n\nระบบจะลบข้อมูลจากฐานข้อมูล และไม่สามารถกู้คืนได้`)) return
    if (!hasSupabaseConfig) {
      const next = communities.filter((item)=>!item.activity_date || item.activity_date >= cutoffDate)
      setCommunities(next); setSelectedId(next[0]?.id ?? null); localStorage.setItem('teacher-community-demo',JSON.stringify(next)); flash(`ลบข้อมูลเก่าก่อนวันที่ ${thaiDate(cutoffDate)} แล้ว`); return
    }
    const storedImages = matches.map((item)=>getStorageObject(item.image_url ?? '')).filter((item):item is {bucket:string;path:string}=>Boolean(item))
    let answer
    try { answer = await withRequestTimeout(supabase!.from('communities').delete().lt('activity_date',cutoffDate),'ลบข้อมูลเก่า') }
    catch (error) { flash(`ลบข้อมูลเก่าไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    if (answer.error) { flash(`ลบข้อมูลเก่าไม่สำเร็จ: ${answer.error.message}`); return }
    await load(false); flash(`ลบข้อมูลเก่าก่อนวันที่ ${thaiDate(cutoffDate)} จากฐานข้อมูลแล้ว กำลังเคลียร์ไฟล์รูป`)
    void Promise.allSettled(storedImages.map((item)=>withRequestTimeout(supabase!.storage.from(item.bucket).remove([item.path]),'ลบไฟล์รูป')))
  }
  function signOut() { if (hasSupabaseConfig) void supabase!.auth.signOut(); localStorage.removeItem(currentAccountKey); setCurrent(null); setView('teacher'); setShowLogin(true); flash('ออกจากระบบแล้ว') }
  async function saveAccount(event:React.FormEvent) {
    event.preventDefault(); if (!accountForm.teacher_name || !accountForm.community_code) { flash('กรุณากรอกชื่อครูและรหัสชุมชน'); return }
    const teacherName = accountForm.teacher_name.trim()
    const communityCode = normalizeCommunityCode(accountForm.community_code)
    const nameDuplicate = accounts.some((x)=>x.id!==editingAccount && normalizeSearch(x.teacher_name)===normalizeSearch(teacherName))
    const codeDuplicate = accounts.some((x)=>x.id!==editingAccount && normalizeSearch(x.community_code)===normalizeSearch(communityCode))
    if (nameDuplicate) { flash('ชื่อครูนี้ถูกใช้แล้ว กรุณาตรวจสอบชื่อหรือแก้ไขรายชื่อเดิม'); return }
    if (codeDuplicate) { flash(`รหัสชุมชน ${communityCode} ถูกใช้แล้ว กรุณาใช้รหัสอื่น`); return }
    if (!hasSupabaseConfig) {
      const item:TeacherAccount = editingAccount ? {...accounts.find((x)=>x.id===editingAccount)!, teacher_name:teacherName, community_code:communityCode} : {id:crypto.randomUUID(),teacher_name:teacherName,community_code:communityCode,role:'teacher'}
      const next=editingAccount?accounts.map((x)=>x.id===item.id?item:x):[...accounts,item]; setAccounts(next); localStorage.setItem('teacher-community-accounts',JSON.stringify(next)); setAccountForm(emptyAccount); setEditingAccount(null); flash(editingAccount?'แก้ไขบัญชีครูแล้ว':'เพิ่มบัญชีครูแล้ว'); return
    }
    let result
    try { result = await withRequestTimeout(supabase!.functions.invoke('admin-accounts',{body:{action:editingAccount?'update':'create',id:editingAccount,teacher_name:teacherName,community_code:communityCode}}),'บันทึกบัญชีครู') }
    catch (error) { flash(`บันทึกไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    const {error} = result
    if (error) { flash(`บันทึกไม่สำเร็จ: ${error.message}`); return }; setAccountForm(emptyAccount); setEditingAccount(null); await load(false); flash('บันทึกสำเร็จแล้ว · บันทึกบัญชีครูแล้ว')
  }
  async function importAccounts(file:File|null) {
    if (!file) return
    setImportingAccounts(true)
    try {
      const workbook = XLSX.read(await file.arrayBuffer(),{type:'array'})
      const imported = extractTeacherRows(workbook).map((item)=>({...item,community_code:normalizeCommunityCode(item.community_code)}))
      const fileCodeOwners = new Map<string,string>()
      for (const item of imported) {
        const codeKey = normalizeSearch(item.community_code)
        const previousName = fileCodeOwners.get(codeKey)
        if (previousName && normalizeSearch(previousName)!==normalizeSearch(item.teacher_name)) {
          throw new Error(`ไฟล์มีรหัสชุมชน ${item.community_code} ซ้ำกัน (${previousName} และ ${item.teacher_name})`)
        }
        fileCodeOwners.set(codeKey,item.teacher_name)
      }
      const existingCodeOwner = imported.find((item)=>accounts.some((account)=>account.role==='teacher' && normalizeSearch(account.community_code)===normalizeSearch(item.community_code) && normalizeSearch(account.teacher_name)!==normalizeSearch(item.teacher_name)))
      if (existingCodeOwner) throw new Error(`รหัสชุมชน ${existingCodeOwner.community_code} ถูกใช้โดยครูคนอื่นแล้ว กรุณาตรวจสอบไฟล์`)
      if (!hasSupabaseConfig) {
        const existingByName = new Map(accounts.map((item)=>[normalizeSearch(item.teacher_name),item]))
        const next = [...accounts]
        for (const item of imported) {
          const existing = existingByName.get(normalizeSearch(item.teacher_name))
          if (existing) {
            const updated = {...existing,community_code:item.community_code}
            const index = next.findIndex((value)=>value.id===existing.id)
            if (index >= 0) next[index] = updated
          } else next.push({id:crypto.randomUUID(),...item,role:'teacher'})
        }
        setAccounts(next); localStorage.setItem('teacher-community-accounts',JSON.stringify(next)); flash(`นำเข้าสำเร็จ ${number(imported.length)} รายชื่อครูแล้ว`); return
      }
      const result = await withRequestTimeout(supabase!.functions.invoke('admin-accounts',{body:{action:'bulk-create',accounts:imported}}),'นำเข้ารายชื่อครู',60000)
      if (result.error) { flash(`นำเข้าไม่สำเร็จ: ${result.error.message}`); return }
      const summary = result.data as {created?:number;updated?:number;failed?:number}
      await load(false)
      flash(`นำเข้าสำเร็จ ${number((summary.created ?? 0)+(summary.updated ?? 0))} รายชื่อ${summary.failed ? ` · ข้าม ${number(summary.failed)} รายการ` : ''}`)
    } catch (error) {
      flash(`นำเข้าไม่สำเร็จ: ${errorText(error,'ไฟล์ไม่ถูกต้องหรือไม่พบคอลัมน์ที่ต้องการ')}`)
    } finally {
      setImportingAccounts(false)
    }
  }
  async function deleteAccount(account:TeacherAccount) {
    if (!window.confirm(`ลบบัญชี ${account.teacher_name} ใช่หรือไม่?`)) return
    if (!hasSupabaseConfig) { const next=accounts.filter((x)=>x.id!==account.id); setAccounts(next); localStorage.setItem('teacher-community-accounts',JSON.stringify(next)); flash('ลบบัญชีครูแล้ว'); return }
    let result
    try { result = await withRequestTimeout(supabase!.functions.invoke('admin-accounts',{body:{action:'delete',id:account.id}}),'ลบบัญชีครู') }
    catch (error) { flash(`ลบไม่สำเร็จ: ${errorText(error,'เชื่อมต่อ Supabase ไม่สำเร็จ')}`); return }
    const {error}=result; if(error){flash(`ลบไม่สำเร็จ: ${error.message}`);return}; await load(false);flash('ลบบัญชีครูแล้ว')
  }
  function editAccount(account:TeacherAccount) { setEditingAccount(account.id); setAccountForm({teacher_name:account.teacher_name,community_code:account.community_code}) }
  function exportExcel() { const sheet=XLSX.utils.json_to_sheet(communities.map((x)=>({'วันที่':thaiDate(x.activity_date),'รหัสชุมชน':x.community_code,'ชื่อชุมชน':x.community_name,'ครูที่ปรึกษา':x.advisor_name,'โรงเรียน':x.school_name,'สถานที่':x.location,'จำนวนสมาชิก':x.member_count,'รายละเอียด':x.description}))); sheet['!cols']=[18,12,30,24,26,20,14,60].map((wch)=>({wch})); const book=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book,sheet,'ชุมชนคุณครู'); XLSX.writeFile(book,'teacher-community.xlsx') }
  async function exportCard() { if(!cardRef.current||!selected)return; const href=await htmlToImage.toPng(cardRef.current,{pixelRatio:2,backgroundColor:'#fff7eb'});const a=document.createElement('a');a.href=href;a.download=`${selected.community_code}-${selected.community_name}.png`;a.click() }
  imagePickerController = {preview:imagePreview,choose:chooseImage,clear:clearImage}
  if (loading) return <main className="app-shell app-bootstrap"><div className="bootstrap-content"><img src="/school-crest.png" alt="ตราโรงเรียน"/><LoaderCircle className="spin" size={24}/><span>กำลังเปิดระบบ…</span></div></main>
  if (connectionError && import.meta.env.PROD && !hasSupabaseConfig) return <ConnectionErrorScreen message={connectionError}/>
  if (showLogin) return <LoginScreen settings={settings} accounts={accounts} teacherSearch={teacherSearch} setTeacherSearch={setTeacherSearch} teacherMatches={teacherMatches} onSubmit={teacherSignIn} onAdmin={()=>{setShowLogin(false);setShowAdminLogin(true)}} notice={connectionError || notice} onDismissNotice={()=>setNotice('')} />
  const AdminPage = AdminMenu
  return <main className="app-shell">
    <header className="topbar safety-topbar"><button className="brand" onClick={()=>setView('teacher')}><span className="school-crest"><img src="/school-crest.png" alt="ตราโรงเรียน"/></span><span>{settings.title}<small>{settings.term} · {settings.year}</small></span></button><nav className="top-actions">{current&&current.role!=='admin'&&<span className="current-teacher-menu">{current.teacher_name}</span>}{!isAdmin&&<button className={view==='teacher'?'nav-button active':'nav-button'} onClick={()=>setView('teacher')}><Pencil size={16}/> ข้อมูลชุมชน</button>}{isAdmin&&<AdminNavigation tab={adminTab} onChange={(next)=>{setAdminTab(next);window.sessionStorage.setItem(adminTabStorageKey,next)}}/>}{current?<button className="login-button logout-icon-button" aria-label="ออกจากระบบ" onClick={signOut}><LogOut size={16}/></button>:<button className="login-button" onClick={()=>setShowLogin(true)}><LogIn size={16}/> ค้นหาชื่อครู</button>}</nav></header>
    {notice&&<div className="notice"><Check size={17}/>{notice}<button onClick={()=>setNotice('')}><X size={16}/></button></div>}
    {view==='teacher'?<section className="teacher-page safety-hero"><div className="teacher-intro"><div className="eyebrow"><Sparkles size={15}/> {settings.school}</div><h1>{current?<>สวัสดี<br/><em>{current.teacher_name}</em></>:<>{settings.title}<br/><em>{settings.term}</em></>}</h1><p>{current?`รหัสชุมชนของคุณคือ ${current.community_code} คุณสามารถบันทึกหรือแก้ไขข้อมูลชุมชนของตนเองได้`:'พิมพ์ชื่อของคุณเพื่อค้นหารายชื่อ แล้วกดเลือกเพื่อเข้าใช้งานได้ทันที'}</p>{current&&<div className="submission-status" role="status"><Check size={17}/><div><b>วันนี้ส่งข้อมูลแล้ว {number(todaySubmissionCount)} ครั้ง</b><span>กดบันทึกซ้ำได้ ระบบจะเก็บทุกครั้งและอัปเดตข้อมูลของวันนี้</span></div></div>}{current&&communities.find((x)=>x.owner_id===current.id&&x.activity_date===today)&&<button className="secondary-action" onClick={()=>openEditor(communities.find((x)=>x.owner_id===current.id&&x.activity_date===today))}><Pencil size={15}/> แก้ไขข้อมูลของวันนี้</button>}</div><form className="community-form" onSubmit={saveCommunity}><div className="form-heading"><div className="form-icon"><BookOpen size={21}/></div><div><h2>{editingId?'แก้ไขข้อมูลชุมนุม':'ข้อมูลชุมนุมของคุณ'}</h2><p>{current?`ข้อมูลประจำวันที่ ${thaiDate(today)} · แก้ไขได้ตลอดวันนี้`:'ค้นหาชื่อครูของคุณก่อนจึงจะบันทึกข้อมูลได้'}</p></div></div><fieldset disabled={!current} className="form-grid"><Field label="วันที่"><input className="readonly-field" type="date" value={form.activity_date} readOnly/></Field><Field label="รหัสชุมชน"><input className="readonly-field" value={form.community_code} readOnly/></Field><Field label="ชื่อชุมนุม" required><input value={form.community_name} onChange={(e)=>setField('community_name',e.target.value)} placeholder="เช่น ร้านค้าสวัสดิการโรงเรียน"/></Field><Field label="ครูที่ปรึกษา" required><input value={form.advisor_name} onChange={(e)=>setField('advisor_name',e.target.value)} placeholder="ชื่อ-นามสกุล"/></Field><Field label="โรงเรียน" required><input value={form.school_name} onChange={(e)=>setField('school_name',e.target.value)} placeholder={settings.school}/></Field><Field label="สถานที่" required><input value={form.location} onChange={(e)=>setField('location',e.target.value)} placeholder="เช่น ห้องคอมพิวเตอร์ 2" required/></Field><Field label="จำนวนที่รับ (อย่างน้อย 22 คน)" required><input type="number" min="22" value={form.member_count} onChange={(e)=>setField('member_count',e.target.value)} placeholder="22" required/></Field><Field label="ลิงก์รูปภาพ (ไม่บังคับ)"><input value={form.image_url} onChange={(e)=>setField('image_url',e.target.value)} placeholder="https://..."/></Field><Field label="รายละเอียดกิจกรรม (ไม่บังคับ · ไม่เกิน 120 ตัวอักษร)" className="span-2"><textarea rows={4} maxLength={TEMPLATE_DETAIL_MAX_LENGTH} value={form.description} onChange={(e)=>setField('description',e.target.value)} placeholder="อธิบายเป้าหมาย กิจกรรม หรือสิ่งที่นักเรียนได้เรียนรู้"/><DescriptionHelp value={form.description}/></Field></fieldset><div className="form-footer"><span><LockKeyhole size={15}/>{current?'ข้อมูลประจำวันที่เลือกไว้แก้ไขได้เฉพาะรายชื่อของคุณ':'ข้อมูลจะถูกปลดล็อกเมื่อเลือกรายชื่อแล้ว'}</span><div>{current?<><button type="button" className="text-button" onClick={()=>{setForm({...emptyForm,activity_date:today,community_code:current?.community_code ?? '',advisor_name:current?.teacher_name ?? '',school_name:settings.school});setEditingId(null)}}>ล้างข้อมูล</button><button className="primary-button" disabled={saving}>{saving?<><LoaderCircle size={17} className="spin"/>กำลังบันทึกข้อมูล...</>:<><ArrowRight size={17}/>{editingId?'บันทึกการแก้ไข':'บันทึกข้อมูลวันนี้'}</>}</button></>:<button type="button" className="primary-button" onClick={()=>setShowLogin(true)}><Search size={17}/> ค้นหาชื่อครู</button>}</div></div></form></section>:<AdminPage accounts={accounts} communities={communities} filtered={filtered} selected={selected} search={search} setSearch={setSearch} loading={loading} accountForm={accountForm} setAccountForm={setAccountForm} editingAccount={editingAccount} onSaveAccount={saveAccount} onEditAccount={editAccount} onDeleteAccount={deleteAccount} onCancelAccount={()=>{setEditingAccount(null);setAccountForm(emptyAccount)}} onImportAccounts={importAccounts} importingAccounts={importingAccounts} onExportExcel={exportExcel} onExportCard={exportCard} onSelected={setSelectedId} cardRef={cardRef} settings={settings} onSaveSettings={updateSettings} onSaveAdminAuth={updateAdminAuth} storageItems={storageItems} storageStats={storageStats} onDeleteImage={deleteCommunityImage} onAdminSaveCommunity={saveCommunityForAdmin} onDeleteDate={deleteCommunitiesByDate} onDeleteBeforeDate={deleteCommunitiesBeforeDate} onDeleteCommunity={deleteCommunity} onLoadDatabaseUsage={loadDatabaseUsage} onRunApiAudit={runApiAudit} adminTab={adminTab}/>}<CreditFooter />
    {showLogin&&<div className="modal-backdrop initial-login-backdrop" onMouseDown={()=>current&&setShowLogin(false)}><section className="login-modal teacher-picker initial-login" onMouseDown={(e)=>e.stopPropagation()}><div className="login-symbol"><Search size={25}/></div><span className="login-kicker">ยินดีต้อนรับ</span><h2>เข้าสู่ระบบ</h2><p>ค้นหาชื่อของคุณ แล้วเลือกชื่อเพื่อเข้าไปกรอกข้อมูลชุมชน</p><label>ชื่อครู<input autoFocus value={teacherSearch} onChange={(e)=>setTeacherSearch(e.target.value)} placeholder="พิมพ์ชื่อครูเพื่อค้นหา..."/></label><div className="teacher-results">{teacherSearch.trim()?teacherMatches.map((teacher)=><button key={teacher.id} onClick={()=>selectTeacher(teacher)}><span>{teacher.teacher_name.slice(0,1)}</span><div><b>{teacher.teacher_name}</b><small><KeyRound size={12}/> รหัสชุมชน {teacher.community_code}</small></div><ArrowRight size={17}/></button>):<div className="search-empty">เริ่มพิมพ์ชื่อ เพื่อค้นหาจาก {number(accounts.filter((item)=>item.role==='teacher').length)} รายชื่อครู</div>}{teacherSearch.trim()&&teacherMatches.length===0&&<div className="search-empty">ไม่พบรายชื่อ ลองพิมพ์คำอื่น หรือแจ้งผู้ดูแลระบบ</div>}</div><button className="admin-entry-button" onClick={()=>{setShowLogin(false);setShowAdminLogin(true)}}><LockKeyhole size={15}/> เข้าสู่ระบบผู้ดูแล</button></section></div>}
    {showAdminLogin&&<div className="modal-backdrop" onMouseDown={()=>{setShowAdminLogin(false);if(!current)setShowLogin(true)}}><section className="login-modal admin-login" onMouseDown={(e)=>e.stopPropagation()}><button className="modal-close" onClick={()=>{setShowAdminLogin(false);if(!current)setShowLogin(true)}}><X size={19}/></button><div className="login-symbol"><LockKeyhole size={25}/></div><h2>เข้าสู่ระบบผู้ดูแล</h2><p>พิมพ์ Username `admin` เพื่อเข้าใช้งานโดยไม่ต้องใช้ Password</p><form onSubmit={adminSignIn}><label>Username<input autoFocus value={adminUsername} onChange={(e)=>setAdminUsername(e.target.value)} placeholder="admin" autoComplete="username" required/></label><button className="primary-button full"><LockKeyhole size={17}/> เข้าสู่ระบบผู้ดูแล</button></form></section></div>}
  </main>
}

type LoginScreenProps = {
  settings:SystemSettings
  accounts:TeacherAccount[]
  teacherSearch:string
  setTeacherSearch:(value:string)=>void
  teacherMatches:TeacherAccount[]
  onSubmit:(event:React.FormEvent)=>void
  onAdmin:()=>void
  notice:string
  onDismissNotice:()=>void
}
function LoginScreen(p:LoginScreenProps){return <main className="login-screen initial-login-backdrop">
  {p.notice&&<div className="login-screen-notice"><Check size={17}/>{p.notice}<button onClick={p.onDismissNotice}><X size={16}/></button></div>}
  <section className="login-layout">
    <section className="login-visual"><div className="login-visual-copy"><span className="login-visual-kicker">พื้นที่เล็ก ๆ สำหรับไอเดียและการเรียนรู้</span><h1>โรงเรียนวิเชียรมาตุ</h1><p>จัดการข้อมูลชุมนุมได้ง่าย ๆ ในไม่กี่ขั้นตอน</p></div><img src="/mascot-welcome.png" alt="น้องชุมนม มาสคอตระบบลงทะเบียนชุมนุม"/></section>
    <section className="login-card initial-login"><div className="login-card-sprout school-logo"><img src="/school-crest.png" alt="ตราโรงเรียน"/></div><h2>ยินดีต้อนรับคุณครู</h2><p>เข้าสู่ระบบเพื่อจัดการข้อมูลชุมชน</p><div className="login-term">{p.settings.term} · {p.settings.year}</div>
      <form className="login-form" onSubmit={p.onSubmit}>
        <label>Username<div className="login-input-wrap"><Search size={21}/><input autoFocus value={p.teacherSearch} onChange={(e)=>p.setTeacherSearch(e.target.value)} placeholder="กรอกชื่อผู้ใช้งาน" autoComplete="username" required/></div></label>
        {p.teacherSearch.trim()&&<div className="username-suggestions">{normalizeSearch(p.teacherSearch)==='admin'?<button type="button" onClick={()=>p.setTeacherSearch('admin')}><span>ผู้</span><b>ผู้ดูแลระบบ</b><small>Username admin · role admin</small></button>:p.teacherMatches.length?p.teacherMatches.map((teacher)=><button type="button" key={teacher.id} onClick={()=>p.setTeacherSearch(teacher.teacher_name)}><span>{teacher.teacher_name.slice(0,1)}</span><b>{teacher.teacher_name}</b><small>{teacher.community_code}</small></button>):<div>ไม่พบรายชื่อที่ตรงกัน</div>}</div>}
        <button className="login-submit" type="submit">เข้าสู่ระบบ <ArrowRight size={23}/></button>
      </form>
    </section>
  </section>
  <CreditFooter />
</main>}

function ConnectionErrorScreen({message}:{message:string}) { return <main className="app-shell app-bootstrap"><div className="bootstrap-content"><img src="/school-crest.png" alt="ตราโรงเรียน"/><Database size={25}/><strong>ยังเชื่อมต่อฐานข้อมูลไม่ได้</strong><span>{message}</span><small>ตั้งค่า Environment Variables ใน Vercel แล้วกด Redeploy อีกครั้ง</small></div></main> }

function CreditFooter(){return <footer className="credit-footer"><div className="credit-footer-line"/><div className="credit-footer-row"><span>ออกแบบและพัฒนาโดย <strong>SASITH WORKS</strong></span><span>รับออกแบบและพัฒนาเว็บไซต์ / เว็บแอปพลิเคชัน</span><a href="mailto:jarungsasitorn@gmail.com">ติดต่อจ้างงาน: jarungsasitorn@gmail.com</a></div></footer>}

function ImageLightbox({src,onClose}:{src:string;onClose:()=>void}) {
  useEffect(()=>{
    function close(event:KeyboardEvent) { if (event.key === 'Escape') onClose() }
    document.addEventListener('keydown',close)
    return ()=>document.removeEventListener('keydown',close)
  },[onClose])
  return <div className="image-lightbox" role="dialog" aria-modal="true" aria-label="ดูรูปภาพขนาดเต็ม" onMouseDown={(event)=>{if(event.target===event.currentTarget)onClose()}}><button type="button" className="image-lightbox-close" onClick={onClose} aria-label="ปิดรูปภาพ"><X size={20}/></button><img src={src} alt="รูปภาพขนาดเต็ม"/></div>
}

function TemplateLightbox({community,teacher,variant,onClose}:{community:Community;teacher:TeacherAccount;variant:TemplateVariant;onClose:()=>void}) {
  useEffect(()=>{
    function close(event:KeyboardEvent) { if (event.key === 'Escape') onClose() }
    document.addEventListener('keydown',close)
    return ()=>document.removeEventListener('keydown',close)
  },[onClose])
  return <div className="image-lightbox template-lightbox" role="dialog" aria-modal="true" aria-label="ดูรูป Template ขนาดเต็ม" onMouseDown={(event)=>{if(event.target===event.currentTarget)onClose()}}><button type="button" className="image-lightbox-close" onClick={onClose} aria-label="ปิดรูป Template"><X size={20}/></button><div className="template-lightbox-card"><TeacherTemplateCard community={community} teacher={teacher} variant={variant}/></div></div>
}

type AdminProps={accounts:TeacherAccount[];communities:Community[];filtered:Community[];selected?:Community;search:string;setSearch:(x:string)=>void;loading:boolean;accountForm:typeof emptyAccount;setAccountForm:(x:typeof emptyAccount)=>void;editingAccount:string|null;onSaveAccount:(e:React.FormEvent)=>void;onEditAccount:(x:TeacherAccount)=>void;onDeleteAccount:(x:TeacherAccount)=>void;onCancelAccount:()=>void;onImportAccounts:(file:File|null)=>Promise<void>;importingAccounts:boolean;onExportExcel:()=>void;onExportCard:()=>void;onSelected:(x:string)=>void;cardRef:React.Ref<HTMLDivElement>;settings:SystemSettings;onSaveSettings:(x:SystemSettings)=>void;onSaveAdminAuth:(x:SystemSettings)=>Promise<void>;storageItems:StorageItem[];storageStats:StorageStats;onDeleteImage:(x:Community)=>void;onAdminSaveCommunity:AdminCommunitySave;onDeleteDate:(date:string)=>Promise<void>;onDeleteBeforeDate:(date:string)=>Promise<void>;onDeleteCommunity:(community:Community)=>Promise<void>;onLoadDatabaseUsage:()=>Promise<DatabaseUsageRow[]>;onRunApiAudit:()=>Promise<ApiAuditItem[]>;adminTab:AdminTab}

function StorageManager(p:{items:StorageItem[];stats:StorageStats;communities:Community[];onDeleteImage:(x:Community)=>void;onDeleteBeforeDate:(date:string)=>Promise<void>;onLoadDatabaseUsage:()=>Promise<DatabaseUsageRow[]>;onRunApiAudit:()=>Promise<ApiAuditItem[]>}) {
  const [selectedImage,setSelectedImage] = useState('')
  const [cutoffDate,setCutoffDate] = useState('')
  const [usageRows,setUsageRows] = useState<DatabaseUsageRow[]>([])
  const [usageLoading,setUsageLoading] = useState(true)
  const [usageError,setUsageError] = useState('')
  const [auditRows,setAuditRows] = useState<ApiAuditItem[]>([])
  const [auditLoading,setAuditLoading] = useState(false)
  const oldRecordCount = cutoffDate ? p.communities.filter((item)=>item.activity_date && item.activity_date < cutoffDate).length : 0
  const databaseBytes = usageRows.filter((row)=>row.object_type==='database table').reduce((total,row)=>total+row.total_bytes,0)
  const storageRow = usageRows.find((row)=>row.object_type==='storage bucket')
  const storageBytes = storageRow?.total_bytes ?? 0
  const totalBytes = databaseBytes + storageBytes
  const refreshUsage = async () => {
    setUsageLoading(true); setUsageError('')
    try { setUsageRows(await p.onLoadDatabaseUsage()) } catch (error) { setUsageError(errorText(error,'โหลดพื้นที่ไม่สำเร็จ')) } finally { setUsageLoading(false) }
  }
  const auditApis = async () => {
    setAuditLoading(true)
    try { setAuditRows(await p.onRunApiAudit()) } finally { setAuditLoading(false) }
  }
  useEffect(()=>{ void refreshUsage() },[])
  return <section className="storage-overview">
    <div className="section-title"><Database size={18}/><div><h3>พื้นที่ข้อมูลและรูปภาพ</h3><p>ตรวจสอบพื้นที่ที่ใช้จริงในฐานข้อมูลและ Supabase Storage ก่อนตัดสินใจลบข้อมูล</p></div></div>
    <div className="storage-stats">
      <div className="storage-stat"><Database size={18}/><b>{usageLoading?'กำลังโหลด…':formatBytes(databaseBytes)}</b><small>ฐานข้อมูล PostgreSQL ทั้งหมด</small></div>
      <div className="storage-stat"><ImageIcon size={18}/><b>{usageLoading?'กำลังโหลด…':formatBytes(storageBytes)}</b><small>ไฟล์ใน Storage จริง</small></div>
      <div className="storage-stat"><Database size={18}/><b>{usageLoading?'กำลังโหลด…':formatBytes(totalBytes)}</b><small>รวมฐานข้อมูล + Storage</small></div>
    </div>
    <div className="storage-capacity"><span>รายงานพื้นที่ล่าสุด</span><b>{usageLoading?'กำลังคำนวณ':usageError?'ตรวจสอบไม่ได้':'เชื่อมต่อแล้ว'}</b><small>รวมขนาดตาราง ดัชนี และไฟล์ Storage ที่อ่านได้จากโปรเจกต์นี้ ไม่ใช่ตัวเลขโควตาแพ็กเกจ</small><button type="button" className="storage-refresh-button" onClick={()=>void refreshUsage()} disabled={usageLoading}><Database size={14}/> รีเฟรชพื้นที่</button></div>
    {usageError&&<div className="storage-note storage-error">{usageError} · ถ้ายังไม่ได้รัน migration `20261004_usage_audit.sql` ให้รันใน Supabase SQL Editor ก่อน</div>}
    <div className="storage-usage-panel"><div className="storage-usage-heading"><div><h4>รายละเอียดทุกส่วนที่ใช้พื้นที่</h4><p>รายการนี้ช่วยดูว่าตารางหรือ Storage ส่วนไหนกินพื้นที่มากที่สุด</p></div><span>{usageRows.length?`${number(usageRows.length)} รายการ`:'ยังไม่มีข้อมูล'}</span></div>{usageRows.length?<div className="storage-usage-table">{[...usageRows].sort((a,b)=>b.total_bytes-a.total_bytes).map((row)=><div className="storage-usage-row" key={`${row.object_type}-${row.object_name}`}><div><b>{row.object_name}</b><small>{row.object_type} · {number(row.row_count)} แถว</small></div><strong>{formatBytes(row.total_bytes)}</strong></div>)}</div>:<div className="storage-empty"><Database size={20}/> กดรีเฟรชพื้นที่เพื่ออ่านข้อมูลจริง</div>}</div>
    <div className="api-audit-panel"><div className="storage-usage-heading"><div><h4>ตรวจสอบเส้น API ทั้งหมด</h4><p>ทดสอบเส้นทางอ่านข้อมูลที่แอปใช้งานและแจ้ง API ที่ช้ากว่า 1 วินาที</p></div><button type="button" className="storage-refresh-button" onClick={()=>void auditApis()} disabled={auditLoading}><Database size={14}/> {auditLoading?'กำลังตรวจสอบ…':'ตรวจสอบ API'}</button></div>{auditRows.length?<div className="api-audit-list">{auditRows.map((row)=><div className={`api-audit-row ${row.status}`} key={row.name}><div><b>{row.name}</b><small>{row.detail}</small></div><strong>{row.status==='not-connected'?'ไม่เชื่อมต่อ':`${number(row.durationMs)} ms`}</strong></div>)}</div>:<div className="storage-empty"><Database size={20}/> กด “ตรวจสอบ API” เพื่อวัดการเชื่อมต่อและความเร็ว</div>}</div>
    <div className="storage-note">ข้อมูลชุมชนและลิงก์รูปภาพถูกเก็บไว้ในฐานข้อมูล การลบข้อมูลเก่าด้านล่างจะลบรายการชุมชนออกจากฐานข้อมูลและลบไฟล์รูปใน Supabase Storage ที่ผูกกับรายการนั้นด้วย ส่วนโควตาแพ็กเกจ เช่น Egress, MAU และจำนวนคำขอ ต้องดูจากหน้า Usage ของ Supabase เพราะไม่เปิดให้ browser client อ่านโดยตรง</div>
    <div className="storage-cleanup"><div><h4>ล้างข้อมูลชุมชนเก่า</h4><p>เลือกวันที่ แล้วระบบจะลบข้อมูลที่บันทึก <b>ก่อนวันที่เลือก</b> ทั้งหมด เหมาะสำหรับเคลียร์ข้อมูลเมื่อใช้งานหลายปี</p></div><div className="storage-cleanup-controls"><label>ลบข้อมูลก่อนวันที่<input type="date" value={cutoffDate} onChange={(event)=>setCutoffDate(event.target.value)}/></label><span className={oldRecordCount?'storage-cleanup-count has-records':'storage-cleanup-count'}>{cutoffDate?`พบ ${number(oldRecordCount)} รายการที่พร้อมลบ`:'กรุณาเลือกวันที่เพื่อดูจำนวนข้อมูล'}</span><button className="danger-button" disabled={!oldRecordCount} onClick={()=>void p.onDeleteBeforeDate(cutoffDate)}><Trash2 size={16}/> ลบข้อมูลเก่าในฐานข้อมูล</button></div></div>
    {p.items.length ? <div className="storage-media-list">{p.items.map((item)=><div className="storage-media-row" key={item.id}><button type="button" className="storage-media-preview" onClick={()=>setSelectedImage(item.image_url ?? '')}><img src={item.image_url} alt={`รูปภาพ${item.community_name}`}/></button><div><b>{item.community_name}</b><small>{item.community_code}</small></div><button className="storage-remove" onClick={()=>p.onDeleteImage(item as Community)}><Trash2 size={15}/> ลบรูป</button></div>)}</div> : <div className="storage-empty"><ImageIcon size={20}/> ยังไม่มีรายการรูปภาพที่บันทึกไว้</div>}
    {selectedImage&&<ImageLightbox src={selectedImage} onClose={()=>setSelectedImage('')}/>} 
  </section>
}
function AdminPage(p:AdminProps){const [draft,setDraft]=useState(p.settings);useEffect(()=>setDraft(p.settings),[p.settings]);return <section className="admin-page safety-admin"><div className="admin-header"><div><div className="eyebrow"><LayoutDashboard size={15}/> ศูนย์ควบคุมผู้ดูแล</div><h1>{p.settings.title}</h1><p>{p.settings.school} · {p.settings.term} · {p.settings.year}</p></div><button className="export-button" onClick={p.onExportExcel}><FileSpreadsheet size={17}/> Export Excel</button></div><div className="admin-metrics"><Metric value={p.communities.length} label="ชุมนุมที่ลงทะเบียน" icon={<BookOpen/>}/><Metric value={p.accounts.filter((x)=>x.role==='teacher').length} label="รายชื่อครู" icon={<Users/>}/><Metric value="กก001–125" label="รหัสชุมนุมจากไฟล์" icon={<KeyRound/>}/></div><div className="admin-columns"><section className="account-manager"><div className="section-title"><UserCog size={18}/><div><h3>จัดการรายชื่อครู</h3><p>ครูค้นหาและเลือกชื่อตัวเองได้ทันที</p></div></div><form className="account-form" onSubmit={p.onSaveAccount}><label>ชื่อครู<input value={p.accountForm.teacher_name} onChange={(e)=>p.setAccountForm({...p.accountForm,teacher_name:e.target.value})} placeholder="ชื่อ-นามสกุลครู"/></label><label>รหัสชุมนุม<input value={p.accountForm.community_code} onChange={(e)=>p.setAccountForm({...p.accountForm,community_code:e.target.value})} placeholder="เช่น กก126"/></label><div><button className="primary-button"><Plus size={16}/>{p.editingAccount?'บันทึกรายชื่อ':'เพิ่มรายชื่อ'}</button>{p.editingAccount&&<button type="button" className="text-button" onClick={p.onCancelAccount}>ยกเลิก</button>}</div></form><div className="account-list">{p.accounts.filter((x)=>x.role==='teacher').map((a)=><div className="account-row" key={a.id}><span className="account-avatar">{a.teacher_name.slice(0,1)}</span><div><b>{a.teacher_name}</b><small><KeyRound size={11}/>{a.community_code}</small></div><button onClick={()=>p.onEditAccount(a)} title="แก้ไข"><Pencil size={15}/></button><button className="danger" onClick={()=>p.onDeleteAccount(a)} title="ลบ"><Trash2 size={15}/></button></div>)}</div></section><div><section className="system-settings"><div className="section-title"><Pencil size={18}/><div><h3>ชื่อระบบและภาคเรียน</h3><p>เปลี่ยนได้ทุกครั้งที่เปิดภาคเรียนใหม่</p></div></div><div className="settings-grid"><label>ชื่อระบบ<input value={draft.title} onChange={(e)=>setDraft({...draft,title:e.target.value})}/></label><label>โรงเรียน<input value={draft.school} onChange={(e)=>setDraft({...draft,school:e.target.value})}/></label><label>ภาคเรียน<input value={draft.term} onChange={(e)=>setDraft({...draft,term:e.target.value})}/></label><label>ปีการศึกษา<input value={draft.year} onChange={(e)=>setDraft({...draft,year:e.target.value})}/></label><label className="span-2">รหัสผู้ดูแล<input type="password" value={draft.adminPassword} onChange={(e)=>setDraft({...draft,adminPassword:e.target.value})}/></label></div><button className="primary-button" onClick={()=>p.onSaveSettings(draft)}><Check size={16}/> บันทึกชื่อระบบ</button></section><section className="community-preview"><div className="list-head"><h3>ตัวอย่างภาพชุมนุม</h3><button className="export-image" onClick={p.onExportCard}><Download size={16}/> PNG</button></div><label className="search"><Search size={17}/><input value={p.search} onChange={(e)=>p.setSearch(e.target.value)} placeholder="ค้นหาชุมนุม..."/></label><div className="admin-workspace"><aside className="community-list"><div className="list-items">{p.loading?<div className="loading"><LoaderCircle className="spin"/>กำลังโหลด...</div>:p.filtered.map((c)=><button key={c.id} className={c.id===p.selected?.id?'community-item selected':'community-item'} onClick={()=>p.onSelected(c.id)}><span className="item-avatar">{c.community_name.slice(0,1)}</span><span><b>{c.community_name}</b><small>{c.community_code} · {c.member_count} สมาชิก</small></span></button>)}</div></aside><div className="preview-pane">{p.selected?<CommunityCard community={p.selected} ref={p.cardRef}/>:<div className="empty">ยังไม่มีข้อมูลชุมนุม</div>}</div></div></section></div></div></section>}
function AdminPageWithStorage(p:AdminProps){return <><StorageManager items={p.storageItems} stats={p.storageStats} communities={p.communities} onDeleteImage={p.onDeleteImage} onDeleteBeforeDate={p.onDeleteBeforeDate} onLoadDatabaseUsage={p.onLoadDatabaseUsage} onRunApiAudit={p.onRunApiAudit}/><AdminPage {...p}/></>}
// Retained for backwards-compatible card export while the visible admin UI uses AdminMenu.
void AdminPageWithStorage

type AdminTab = 'entry'|'template'|'storage'|'teachers'|'settings'|'admin-auth'

const adminTabs:{id:AdminTab;label:string;icon:React.ReactNode}[] = [
  {id:'entry',label:'กรอกข้อมูลชุมชน',icon:<Pencil size={17}/>},
  {id:'template',label:'Template',icon:<FileSpreadsheet size={17}/>},
  {id:'storage',label:'พื้นที่ข้อมูลและรูปภาพ',icon:<Database size={17}/>},
  {id:'teachers',label:'จัดการรายชื่อครู',icon:<Users size={17}/>},
  {id:'settings',label:'ชื่อระบบและภาคเรียน',icon:<Pencil size={17}/>},
  {id:'admin-auth',label:'บัญชีผู้ดูแล',icon:<LockKeyhole size={17}/>},
]

function AdminNavigation({tab,onChange}:{tab:AdminTab;onChange:(tab:AdminTab)=>void}) {
  const [open,setOpen] = useState(false)
  const rootRef = useRef<HTMLDivElement>(null)
  useEffect(()=>{
    function close(event:MouseEvent) {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown',close)
    return ()=>document.removeEventListener('mousedown',close)
  },[])
  function choose(next:AdminTab) { onChange(next); setOpen(false) }
  return <div className={open?'admin-navigation is-open':'admin-navigation'} ref={rootRef}>
    <button type="button" className="admin-navigation-trigger" aria-expanded={open} aria-haspopup="true" onClick={()=>setOpen((current)=>!current)}><Menu size={17}/><span>เมนูผู้ดูแล</span><ChevronDown size={15}/></button>
    <nav className="admin-navigation-menu" aria-label="เมนูผู้ดูแล">{adminTabs.map((item)=><button type="button" key={item.id} className={tab===item.id?'admin-tab active':'admin-tab'} onClick={()=>choose(item.id)}>{item.icon}{item.label}</button>)}</nav>
  </div>
}

function AdminMenu(p:AdminProps) {
  return <section className="admin-page safety-admin admin-menu-page">
    {!hasSupabaseConfig&&<div className="storage-note"><b>กำลังใช้งานโหมดตัวอย่าง</b> รายชื่อและข้อมูลที่เห็นยังมาจากเครื่องนี้ ไม่ใช่ Supabase · เพิ่มค่า `VITE_SUPABASE_URL` และ `VITE_SUPABASE_ANON_KEY` ใน `.env.local` แล้วรีสตาร์ตเว็บเพื่อใช้ข้อมูลจริง</div>}
    {p.adminTab==='entry'&&<AdminCommunityEntryPanel accounts={p.accounts} communities={p.communities} settings={p.settings} onSave={p.onAdminSaveCommunity}/>}
    {p.adminTab==='template'&&<TemplatePanel communities={p.communities} accounts={p.accounts} onDeleteDate={p.onDeleteDate} onDeleteCommunity={p.onDeleteCommunity}/>}
    {p.adminTab==='storage'&&<StorageManager items={p.storageItems} stats={p.storageStats} communities={p.communities} onDeleteImage={p.onDeleteImage} onDeleteBeforeDate={p.onDeleteBeforeDate} onLoadDatabaseUsage={p.onLoadDatabaseUsage} onRunApiAudit={p.onRunApiAudit}/>}
    {p.adminTab==='teachers'&&<TeacherManagerPanel {...p}/>}
    {p.adminTab==='settings'&&<SystemSettingsPanel settings={p.settings} onSaveSettings={p.onSaveSettings}/>}
    {p.adminTab==='admin-auth'&&<AdminAuthPanel settings={p.settings} onSaveAuth={p.onSaveAdminAuth}/>}
  </section>
}

function AdminCommunityEntryPanel({accounts,communities,settings,onSave}:{accounts:TeacherAccount[];communities:Community[];settings:SystemSettings;onSave:AdminCommunitySave}) {
  const teachers = accounts.filter((account)=>account.role==='teacher')
  const [teacherId,setTeacherId] = useState(teachers[0]?.id ?? '')
  const [form,setForm] = useState<CommunityForm>({...emptyForm,activity_date:todayInBangkok()})
  const [imageFile,setImageFile] = useState<File|null>(null)
  const [imagePreview,setImagePreview] = useState('')
  const [imageNotice,setImageNotice] = useState('')
  const [showImage,setShowImage] = useState(false)
  const teacher = teachers.find((account)=>account.id===teacherId)
  const record = teacher ? communities.find((item)=>item.owner_id===teacher.id && item.activity_date===form.activity_date) : undefined
  useEffect(()=>{ if (!teacherId && teachers[0]) setTeacherId(teachers[0].id) },[teacherId,teachers])
  useEffect(()=>{
    if (!teacher) return
    const existing = communities.find((item)=>item.owner_id===teacher.id && item.activity_date===form.activity_date)
    if (existing) {
      setForm({activity_date:existing.activity_date,community_code:existing.community_code,community_name:existing.community_name,advisor_name:existing.advisor_name,school_name:existing.school_name,location:existing.location,member_count:String(existing.member_count),description:limitTemplateDetail(existing.description),image_url:existing.image_url ?? ''})
      setImagePreview(existing.image_url ?? '')
    } else {
      setForm({...emptyForm,activity_date:form.activity_date,community_code:teacher.community_code,advisor_name:teacher.teacher_name,school_name:settings.school})
      setImagePreview('')
    }
    setImageFile(null); setImageNotice('')
  },[teacherId,communities,settings.school,form.activity_date])
  function setField(key:keyof CommunityForm,value:string) { setForm((current)=>({...current,[key]:key==='description'?limitTemplateDetail(value):value})) }
  async function chooseImage(file:File|null) {
    if (!file) return
    if (!file.type.startsWith('image/')) { setImageNotice('กรุณาเลือกไฟล์รูปภาพเท่านั้น'); return }
    if (file.size > 12*1024*1024) { setImageNotice('รูปภาพต้นฉบับต้องมีขนาดไม่เกิน 12 MB'); return }
    try {
      const compact = await compressImage(file)
      if (imagePreview.startsWith('blob:')) URL.revokeObjectURL(imagePreview)
      setImageFile(compact); setImagePreview(URL.createObjectURL(compact)); setImageNotice(`ย่อรูปเหลือ ${formatBytes(compact.size)} แล้ว`)
    } catch (error) { setImageNotice(error instanceof Error ? error.message : 'บีบอัดรูปไม่สำเร็จ') }
  }
  function clearImage() { if (imagePreview.startsWith('blob:')) URL.revokeObjectURL(imagePreview); setImageFile(null); setImagePreview(''); setField('image_url',''); setImageNotice('') }
  if (!teachers.length) return <section className="admin-panel admin-entry-panel"><div className="template-empty"><Users size={22}/><b>ยังไม่มีรายชื่อคุณครู</b><span>เพิ่มรายชื่อคุณครูในเมนูจัดการรายชื่อครูก่อน</span></div></section>
  return <section className="admin-panel admin-entry-panel">
    <div className="section-title"><Pencil size={20}/><div><h3>กรอกข้อมูลชุมชนแทนคุณครู</h3><p>เลือกชื่อคุณครูเพื่อกรอกหรือแก้ไขข้อมูลชุมชนของคนนั้น</p></div></div>
    <div className="admin-entry-selector"><label>เลือกคุณครู<TeacherCombobox teachers={teachers} value={teacherId} onChange={setTeacherId}/></label><span className={record?'has-record':''}>{record?'มีข้อมูลแล้ว สามารถแก้ไขได้':'ยังไม่มีข้อมูลของวันนี้'}</span></div>
    <form className="admin-entry-form" onSubmit={(event)=>{event.preventDefault(); if(teacher) void onSave(teacher,form,imageFile,imagePreview)}}>
      <fieldset className="admin-entry-grid">
        <label className="field"><span>วันที่</span><input type="date" value={form.activity_date} onChange={(event)=>setField('activity_date',event.target.value)}/></label>
        <label className="field"><span>รหัสชุมชน</span><input className="readonly-field" value={teacher?.community_code ?? ''} readOnly/></label>
        <label className="field"><span>ชื่อชุมนุม<i>*</i></span><input value={form.community_name} onChange={(event)=>setField('community_name',event.target.value)} placeholder="เช่น ร้านค้าสวัสดิการโรงเรียน" required/></label>
        <label className="field"><span>ครูที่ปรึกษา<i>*</i></span><input className="readonly-field" value={form.advisor_name || teacher?.teacher_name || ''} readOnly required/></label>
        <label className="field"><span>โรงเรียน<i>*</i></span><input className="readonly-field" value={form.school_name || settings.school} readOnly required/></label>
        <label className="field"><span>สถานที่<i>*</i></span><input value={form.location} onChange={(event)=>setField('location',event.target.value)} placeholder="เช่น ห้องคอมพิวเตอร์ 2" required/></label>
        <label className="field"><span>จำนวนที่รับ (อย่างน้อย 22 คน)<i>*</i></span><input type="number" min="22" value={form.member_count} onChange={(event)=>setField('member_count',event.target.value)} placeholder="22" required/></label>
        <div className="field image-field"><span>รูปภาพคุณครูประจำชุมชน<i>*</i></span><div className="image-picker"><input id="admin-community-image-file" type="file" accept="image/*" onChange={(event)=>void chooseImage(event.target.files?.[0] ?? null)}/><input id="admin-community-image-camera" type="file" accept="image/*" capture="environment" onChange={(event)=>void chooseImage(event.target.files?.[0] ?? null)}/><div className="image-picker-actions"><label htmlFor="admin-community-image-file"><ImagePlus size={16}/> เลือกรูป</label><label htmlFor="admin-community-image-camera"><Camera size={16}/> ถ่ายรูป</label></div>{imagePreview&&<div className="image-preview"><button type="button" className="image-preview-trigger" onClick={()=>setShowImage(true)} aria-label="ดูรูปภาพขนาดเต็ม"><img src={imagePreview} alt="ตัวอย่างรูปคุณครูประจำชุมชน"/></button><button type="button" onClick={clearImage}><X size={14}/> ลบรูป</button></div>}<small>{imageNotice || 'คลิกรูปเพื่อดูขนาดเต็ม · ระบบจะย่อรูปไม่เกิน 500 KB และเก็บเมื่อกดบันทึกข้อมูลเท่านั้น'}</small></div></div>
        <label className="field span-2"><span>รายละเอียดกิจกรรม (ไม่บังคับ · ไม่เกิน 120 ตัวอักษร)</span><textarea rows={4} maxLength={TEMPLATE_DETAIL_MAX_LENGTH} value={form.description} onChange={(event)=>setField('description',event.target.value)} placeholder="อธิบายเป้าหมาย กิจกรรม หรือสิ่งที่นักเรียนได้เรียนรู้"/><DescriptionHelp value={form.description}/></label>
      </fieldset>
      <div className="admin-entry-footer"><span>{record?'กำลังแก้ไขข้อมูลของคุณครูคนนี้':'กำลังกรอกข้อมูลแทนคุณครูคนนี้'}</span><button className="primary-button"><Check size={16}/> {record?'บันทึกการแก้ไข':'บันทึกข้อมูล'}</button></div>
    </form>
    {showImage&&imagePreview&&<ImageLightbox src={imagePreview} onClose={()=>setShowImage(false)}/>} 
  </section>
}

function TeacherCombobox({teachers,value,onChange,showAll=false}:{teachers:TeacherAccount[];value:string;onChange:(value:string)=>void;showAll?:boolean}) {
  const [open,setOpen] = useState(false)
  const [query,setQuery] = useState('')
  const rootRef = useRef<HTMLDivElement>(null)
  const selected = teachers.find((teacher)=>teacher.id===value)
  const results = useMemo(()=>{
    const needle = normalizeSearch(query)
    const matches = needle ? teachers.filter((teacher)=>[teacher.teacher_name,teacher.community_code,teacher.login_email ?? ''].some((value)=>normalizeSearch(value).includes(needle))) : teachers
    return matches.slice(0,60)
  },[teachers,query])

  useEffect(()=>{
    function close(event:MouseEvent) {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown',close)
    return ()=>document.removeEventListener('mousedown',close)
  },[])

  function choose(id:string) {
    onChange(id)
    setQuery('')
    setOpen(false)
  }

  return <div className="teacher-combobox" ref={rootRef}>
    <button type="button" className={open?'teacher-combobox-trigger is-open':'teacher-combobox-trigger'} aria-haspopup="listbox" aria-expanded={open} onClick={()=>setOpen((current)=>!current)}>
      <span className={selected||value==='all'?'teacher-combobox-value':'teacher-combobox-placeholder'}>{value==='all'?'คุณครูทั้งหมด':selected?selected.teacher_name:'เลือกชื่อคุณครู'}</span>
      {selected&&<small>{selected.community_code}</small>}
      {value==='all'&&<small>ทุกคน</small>}
      <ChevronDown size={18}/>
    </button>
    {open&&<div className="teacher-combobox-menu" role="listbox" aria-label="รายชื่อคุณครู">
      <div className="teacher-combobox-search"><Search size={16}/><input autoFocus value={query} onChange={(event)=>setQuery(event.target.value)} onKeyDown={(event)=>{if(event.key==='Escape')setOpen(false);if(event.key==='Enter'&&results[0])choose(results[0].id)}} placeholder="ค้นหาชื่อหรือรหัสชุมชน..."/></div>
      <div className="teacher-combobox-options">
        {showAll&&<button type="button" role="option" aria-selected={value==='all'} className={value==='all'?'is-selected':''} onClick={()=>choose('all')}><span aria-hidden="true">✓</span><div><b>คุณครูทั้งหมด</b><small>แสดงข้อมูลทุกคน</small></div>{value==='all'&&<Check size={16}/>}</button>}
        {results.length?results.map((teacher)=><button type="button" role="option" aria-selected={teacher.id===value} className={teacher.id===value?'is-selected':''} key={teacher.id} onClick={()=>choose(teacher.id)}><span>{teacher.teacher_name.slice(0,1)}</span><div><b>{teacher.teacher_name}</b><small>{teacher.community_code}</small></div>{teacher.id===value&&<Check size={16}/>}</button>):<div className="teacher-combobox-empty">ไม่พบรายชื่อ “{query.trim()}” ลองใช้ชื่อภาษาไทยหรือรหัสชุมชน</div>}
      </div>
    </div>}
  </div>
}

function TemplatePanel({communities,accounts,onDeleteDate,onDeleteCommunity}:{communities:Community[];accounts:TeacherAccount[];onDeleteDate:(date:string)=>Promise<void>;onDeleteCommunity:(community:Community)=>Promise<void>}) {
  const [selectedDate,setSelectedDate] = useState(todayInBangkok())
  const [templateVariant,setTemplateVariant] = useState<TemplateVariant>('html')
  const teachers = useMemo(()=>accounts.filter((account)=>account.role==='teacher'),[accounts])
  const [selectedTeacherId,setSelectedTeacherId] = useState('all')
  const [selectedIds,setSelectedIds] = useState<string[]>([])
  const [page,setPage] = useState(1)
  const pageSize = 20
  const daily = useMemo(()=>communities.filter((item)=>item.activity_date===selectedDate),[communities,selectedDate])
  const selectedTeacher = teachers.find((teacher)=>teacher.id===selectedTeacherId)
  const visibleDaily = useMemo(()=>selectedTeacherId==='all' ? daily : daily.filter((item)=>selectedTeacher && (item.owner_id===selectedTeacher.id || item.community_code===selectedTeacher.community_code)),[daily,selectedTeacher,selectedTeacherId])
  const selectedItems = useMemo(()=>visibleDaily.filter((item)=>selectedIds.includes(item.id)),[selectedIds,visibleDaily])
  const exportRefs = useRef<Record<string,HTMLDivElement|null>>({})
  const pageCount = Math.max(1,Math.ceil(visibleDaily.length/pageSize))
  const currentPage = Math.min(page,pageCount)
  const pagedDaily = visibleDaily.slice((currentPage-1)*pageSize,currentPage*pageSize)
  const allVisibleSelected = visibleDaily.length > 0 && visibleDaily.every((item)=>selectedIds.includes(item.id))
  const [previewCommunity,setPreviewCommunity] = useState<Community|null>(null)
  useEffect(()=>{
    if (selectedTeacherId !== 'all' && selectedTeacherId && !selectedTeacher) setSelectedTeacherId('all')
  },[selectedTeacherId,selectedTeacher,teachers])
  useEffect(()=>{
    setSelectedIds((current)=>current.filter((id)=>visibleDaily.some((item)=>item.id===id)))
  },[visibleDaily])
  const teacherName = (item:Community) => accounts.find((account)=>account.id===item.owner_id)?.teacher_name ?? item.advisor_name
  const templateTeacher = (item:Community):TeacherAccount => accounts.find((account)=>account.id===item.owner_id) ?? accounts.find((account)=>account.community_code===item.community_code) ?? {id:`template-${item.id}`,teacher_name:item.advisor_name,community_code:item.community_code,role:'teacher'}
  async function exportTemplate() {
    for (const item of selectedItems) {
      const node = exportRefs.current[item.id]
      if (!node) continue
      const dataUrl = await htmlToImage.toPng(node,{pixelRatio:2,cacheBust:true,backgroundColor:'#fff7eb'})
      const link = document.createElement('a')
      link.download = `template-${item.community_code}-${selectedDate}.png`
      link.href = dataUrl
      link.click()
      await new Promise((resolve)=>setTimeout(resolve,180))
    }
  }
  function exportDaily() {
    const sheet = XLSX.utils.json_to_sheet(selectedItems.map((item)=>({'วันที่':thaiDate(selectedDate),'ชื่อครู':teacherName(item),'รหัสชุมชน':item.community_code,'ชื่อชุมนุม':item.community_name,'โรงเรียน':item.school_name,'สถานที่':item.location,'จำนวนสมาชิก':item.member_count,'รายละเอียดกิจกรรม':item.description,'ลิงก์รูปภาพ':item.image_url ?? ''})))
    sheet['!cols'] = [18,26,14,30,26,20,14,60,42].map((wch)=>({wch}))
    const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book,sheet,'ข้อมูลรายวัน')
    XLSX.writeFile(book,`teacher-community-${selectedDate}-${selectedTeacherId==='all'?'all':selectedTeacher?.community_code ?? 'teacher'}.xlsx`)
  }
  return <section className="admin-panel template-panel">
    <div className="section-title"><FileSpreadsheet size={20}/><div><h3>Template ภาพชุมนุม</h3><p>เลือกวันที่และชื่อคุณครู เพื่อสร้างภาพรูปแบบเดียวกันสำหรับใช้งานหรือดาวน์โหลด</p></div></div>
    <div className="template-toolbar template-toolbar-template"><label>วันที่บันทึก<input type="date" value={selectedDate} onChange={(event)=>{setSelectedDate(event.target.value);setSelectedIds([]);setPage(1)}}/></label><label className="template-teacher-filter">คุณครู<TeacherCombobox teachers={teachers} value={selectedTeacherId} showAll onChange={(next)=>{setSelectedTeacherId(next);setSelectedIds([]);setPage(1)}}/></label><label className="template-variant-picker">รูปแบบ Template<select value={templateVariant} onChange={(event)=>setTemplateVariant(event.target.value as TemplateVariant)}>{Object.entries(templateVariants).map(([value,variant])=><option key={value} value={value}>{variant.label}</option>)}</select></label><div className="template-toolbar-actions"><span className="template-selection-summary">เลือกแล้ว {number(selectedItems.length)} รายการ</span><button className="primary-button" disabled={!selectedItems.length} onClick={()=>void exportTemplate()}><Download size={16}/> ดาวน์โหลด PNG</button><button className="secondary-button" disabled={!selectedItems.length} onClick={exportDaily}><FileSpreadsheet size={16}/> Export Xlsx</button><button className="danger-button" disabled={!daily.length} onClick={()=>void onDeleteDate(selectedDate)}><Trash2 size={16}/> ลบข้อมูลวันนี้</button></div></div>
    <div className="template-export-staging" aria-hidden="true">{selectedItems.map((item)=><TeacherTemplateCard key={item.id} community={item} teacher={templateTeacher(item)} variant={templateVariant} ref={(node)=>{exportRefs.current[item.id]=node}}/>)}</div>
    {visibleDaily.length ? <>
      <div className="template-day-summary"><b>ตารางข้อมูล Template ของวันที่เลือก</b><span>{number(visibleDaily.length)} รายการ · {selectedTeacherId==='all'?'แสดงข้อมูลคุณครูทั้งหมด':'แสดงข้อมูลของคุณครูที่เลือก'}</span></div>
      <div className="template-table-wrap"><table className="template-table"><thead><tr><th className="template-select-column"><input type="checkbox" checked={allVisibleSelected} onChange={()=>setSelectedIds(allVisibleSelected?[]:visibleDaily.map((item)=>item.id))} aria-label="เลือกข้อมูลทั้งหมด"/></th><th className="template-image-column">รูป Template</th><th>ครูผู้บันทึก</th><th>รหัสชุมชน</th><th>ชื่อชุมนุม</th><th>โรงเรียน</th><th>จำนวนที่รับ</th><th>สถานที่</th><th className="template-description-column">รายละเอียดกิจกรรม</th><th className="template-actions-column">จัดการ</th></tr></thead><tbody>{pagedDaily.map((item)=><tr key={item.id}><td className="template-select-column"><input type="checkbox" checked={selectedIds.includes(item.id)} onChange={()=>setSelectedIds((current)=>current.includes(item.id)?current.filter((id)=>id!==item.id):[...current,item.id])} aria-label={`เลือก ${item.community_name}`}/></td><td className="template-image-column"><button type="button" className="template-card-mini-button" onClick={()=>setPreviewCommunity(item)} title="กดเพื่อดูรูป Template ขนาดใหญ่"><div className="template-card-mini"><TeacherTemplateCard community={item} teacher={templateTeacher(item)} variant={templateVariant} ref={undefined}/></div></button></td><td>{teacherName(item)}</td><td>{item.community_code}</td><td>{item.community_name}</td><td>{item.school_name}</td><td>{number(item.member_count)}</td><td>{item.location||'-'}</td><td className="template-description-cell">{item.description||'-'}</td><td className="template-actions-cell"><button type="button" className="template-delete-button" onClick={()=>void onDeleteCommunity(item)} title="ลบข้อมูลแถวนี้"><Trash2 size={15}/> ลบ</button></td></tr>)}</tbody></table></div>
      {pageCount>1&&<TemplatePagination page={currentPage} pageCount={pageCount} onPageChange={setPage}/>} 
    </> : <div className="template-empty"><FileSpreadsheet size={22}/><b>ยังไม่มีข้อมูลในวันที่เลือก</b><span>ลองเลือกวันอื่น หรือรอคุณครูบันทึกข้อมูลชุมชน</span></div>}
    {previewCommunity&&<TemplateLightbox community={previewCommunity} teacher={templateTeacher(previewCommunity)} variant={templateVariant} onClose={()=>setPreviewCommunity(null)}/>}
  </section>
}

function fitTemplateText(value:string,maxLength:number) { return Array.from(value).slice(0,maxLength).join('') }

function fitPosterText(field:HTMLElement,max:number,min:number) {
  const text = field.firstElementChild as HTMLElement | null
  if (!text) return
  const fits = (size:number) => { field.style.fontSize = `${size}px`; return text.scrollWidth <= text.clientWidth + 1 && field.scrollHeight <= field.clientHeight + 1 }
  if (fits(max)) return
  let low = min, high = max
  while (high - low > .5) { const middle = (low + high) / 2; if (fits(middle)) low = middle; else high = middle }
  field.style.fontSize = `${low}px`
}

function PosterText({name,field,children}:{name:TemplateFieldKey;field:PosterField;children:React.ReactNode}) {
  const fieldRef = useRef<HTMLDivElement>(null)
  useLayoutEffect(()=>{
    const element = fieldRef.current
    if (!element) return
    let active = true
    fitPosterText(element,field.max,field.min)
    // Web fonts load lazily and change text width, so measure again whenever one finishes loading.
    const refit = () => { if (active) fitPosterText(element,field.max,field.min) }
    const style = window.getComputedStyle(element)
    void Promise.all([document.fonts?.load(`${style.fontWeight} ${field.max}px ${style.fontFamily}`),document.fonts?.ready]).catch(()=>undefined).then(refit)
    document.fonts?.addEventListener('loadingdone',refit)
    return ()=>{ active = false; document.fonts?.removeEventListener('loadingdone',refit) }
  },[children,field])
  return <div ref={fieldRef} className={`poster-field poster-field-${name} ${field.className ?? ''}`} style={{...field.box,fontSize:field.max}}><span>{children}</span></div>
}

function HtmlPosterBackdrop() {
  return <>
    <div className="html-poster-sky"/><div className="html-poster-blob blob-one"/><div className="html-poster-blob blob-two"/><div className="html-poster-blob blob-three"/>
    <img className="html-poster-crest" src="/school-crest.png" alt="ตราโรงเรียน"/>
    <div className="html-poster-title"><h1><span>กิจกรรม</span><span>ชุมนุม</span></h1><strong>โรงเรียนวิเชียรมาตุ</strong></div>
    <div className="html-poster-photo-frame"/>
    <div className="html-poster-caption">ครูที่ปรึกษาชุมนุม</div>
    {Object.values(htmlPosterCards).map(({label,tone,card})=><div key={label} className={`html-poster-card tone-${tone}`} style={card}><span>{label}</span></div>)}
    <div className="html-poster-footer">ค้นหาความชอบ เติบโตไปด้วยกัน</div>
  </>
}

function PosterTeacherTemplateCard({community,teacher,variant,ref}:{community:Community;teacher:TeacherAccount;variant:TemplateVariant;ref?:React.Ref<HTMLDivElement>}) {
  const cardRef = useRef<HTMLDivElement|null>(null)
  const [scale,setScale] = useState(1)
  useLayoutEffect(()=>{
    const element = cardRef.current
    if (!element) return
    const update = () => { if (element.clientWidth) setScale(element.clientWidth / POSTER_SIZE) }
    update()
    const observer = new ResizeObserver(update)
    observer.observe(element)
    return ()=>observer.disconnect()
  },[])
  const setRefs = (node:HTMLDivElement|null) => { cardRef.current = node; if (typeof ref === 'function') ref(node); else if (ref) (ref as React.RefObject<HTMLDivElement|null>).current = node }
  const template = templateVariants[variant]
  const fields = template.fields
  const advisor = community.advisor_name || teacher.teacher_name
  return <article className={`teacher-template-card original-template-card poster-template-card template-variant-${variant}`} ref={setRefs} aria-label={`Template ชุมนุม ${community.community_name}`}>
    <div className="poster-stage" style={{width:POSTER_SIZE,height:POSTER_SIZE,transform:`scale(${scale})`}}>
      {template.background?<img className="poster-background" src={template.background} alt="" aria-hidden="true"/>:<HtmlPosterBackdrop/>}
      <div className="poster-photo" style={template.photo}>{community.image_url?<img src={community.image_url} alt={`รูปคุณครู ${advisor}`}/>:<div className="poster-photo-empty"><ImagePlus size={64}/><span>รูปภาพคุณครู<br/>ประจำชุมนุม</span></div>}</div>
      <PosterText name="code" field={fields.code}>{community.community_code || '-'}</PosterText>
      <PosterText name="name" field={fields.name}>{community.community_name || '-'}</PosterText>
      <PosterText name="advisor" field={fields.advisor}>{advisor}</PosterText>
      <PosterText name="location" field={fields.location}>{community.location || '-'}</PosterText>
      <PosterText name="members" field={fields.members}>{community.member_count ? `${number(community.member_count)} คน` : '-'}</PosterText>
      <PosterText name="description" field={fields.description}>{fitTemplateText(community.description || '-',TEMPLATE_DETAIL_MAX_LENGTH)}</PosterText>
    </div>
  </article>
}

function TeacherTemplateCard({community,teacher,variant='html',ref}:{community:Community;teacher:TeacherAccount;variant?:TemplateVariant;ref?:React.Ref<HTMLDivElement>}) {
  return <PosterTeacherTemplateCard community={community} teacher={teacher} variant={variant} ref={ref}/>
}

function DescriptionHelp({value}:{value:string}) {
  const length = Array.from(value).length
  const tooLong = length > TEMPLATE_DETAIL_RECOMMENDED_LENGTH
  return <small className={`field-help${tooLong?' field-help-warning':''}`}>{length}/{TEMPLATE_DETAIL_MAX_LENGTH} ตัวอักษร · {tooLong?`ยาวเกิน ${TEMPLATE_DETAIL_RECOMMENDED_LENGTH} ตัวอักษร ตัวหนังสือใน Template จะเล็กลง`:`แนะนำไม่เกิน ${TEMPLATE_DETAIL_RECOMMENDED_LENGTH} ตัวอักษร (ประมาณ 2 ประโยคสั้น) เพื่อให้ Template สวยที่สุด`}</small>
}

function TemplatePagination({page,pageCount,onPageChange}:{page:number;pageCount:number;onPageChange:(page:number)=>void}) {
  const items:(number|'ellipsis')[] = pageCount<=5
    ? Array.from({length:pageCount},(_,index)=>index+1)
    : page<=3
      ? [1,2,3,4,'ellipsis',pageCount]
      : page>=pageCount-2
        ? [1,'ellipsis',pageCount-3,pageCount-2,pageCount-1,pageCount]
        : [1,'ellipsis',page-1,page,page+1,'ellipsis',pageCount]
  return <Pagination>
    <PaginationContent>
      <PaginationItem><PaginationPrevious disabled={page===1} onClick={()=>onPageChange(Math.max(1,page-1))}/></PaginationItem>
      {items.map((item,index)=><PaginationItem key={`${item}-${index}`}>{item==='ellipsis'?<PaginationEllipsis/>:<PaginationLink isActive={item===page} onClick={()=>onPageChange(item)}>{item}</PaginationLink>}</PaginationItem>)}
      <PaginationItem><PaginationNext disabled={page===pageCount} onClick={()=>onPageChange(Math.min(pageCount,page+1))}/></PaginationItem>
    </PaginationContent>
  </Pagination>
}

function TeacherManagerPanel(p:AdminProps) {
  const teachers = p.accounts.filter((item)=>item.role==='teacher')
  return <section className="account-manager admin-panel">
    <div className="section-title"><UserCog size={18}/><div><h3>จัดการรายชื่อครู</h3><p>เพิ่ม แก้ไข หรือลบรายชื่อครูและรหัสชุมชน</p></div></div>
    <form className="account-form" onSubmit={p.onSaveAccount}>
      <label>ชื่อครู<input value={p.accountForm.teacher_name} onChange={(event)=>p.setAccountForm({...p.accountForm,teacher_name:event.target.value})} placeholder="ชื่อ-นามสกุลครู"/></label>
      <label>รหัสชุมชน<input value={p.accountForm.community_code} onChange={(event)=>p.setAccountForm({...p.accountForm,community_code:event.target.value})} placeholder="เช่น กก126"/></label>
      <div className="account-form-actions">
        <button className="primary-button"><Plus size={16}/>{p.editingAccount?'บันทึกรายชื่อ':'เพิ่มรายชื่อ'}</button>
        <label className="secondary-button import-accounts-button"><FileSpreadsheet size={16}/>{p.importingAccounts?'กำลังนำเข้า...':'นำเข้าไฟล์'}<input type="file" accept=".xlsx,.xls,.csv" disabled={p.importingAccounts} onChange={(event)=>{void p.onImportAccounts(event.target.files?.[0] ?? null);event.currentTarget.value=''}}/></label>
        {p.editingAccount&&<button type="button" className="text-button" onClick={p.onCancelAccount}>ยกเลิก</button>}
      </div>
    </form>
    <p className="import-accounts-help">นำเข้าไฟล์ Excel ที่มีคอลัมน์ “ครูที่ปรึกษา” และ “รหัสชุมนุม” ระบบจะสร้างหรืออัปเดตรายชื่อให้เป็นชุดเดียว</p>
    <div className="account-list">{teachers.map((account)=><div className="account-row" key={account.id}><span className="account-avatar">{account.teacher_name.slice(0,1)}</span><div><b>{account.teacher_name}</b><small><KeyRound size={11}/>{account.community_code}</small></div><button onClick={()=>p.onEditAccount(account)} title="แก้ไข"><Pencil size={15}/></button><button className="danger" onClick={()=>p.onDeleteAccount(account)} title="ลบ"><Trash2 size={15}/></button></div>)}</div>
  </section>
}

function AdminAuthPanel({settings,onSaveAuth}:{settings:SystemSettings;onSaveAuth:(settings:SystemSettings)=>Promise<void>}) {
  const [draft,setDraft] = useState(settings)
  useEffect(()=>setDraft(settings),[settings])
  return <section className="admin-auth-panel admin-panel"><div className="section-title"><LockKeyhole size={18}/><div><h3>บัญชีผู้ดูแล</h3><p>กำหนด Username สำหรับเข้าสู่ระบบโดยไม่ต้องใช้ Password</p></div></div><form className="admin-auth-form" onSubmit={(event)=>{event.preventDefault(); if(!draft.adminUsername.trim()) return; void onSaveAuth({...draft,adminUsername:draft.adminUsername.trim()})}}><label>Username<input value={draft.adminUsername} onChange={(event)=>setDraft({...draft,adminUsername:event.target.value})} placeholder="admin" autoComplete="username" required/></label><p className="admin-auth-note">ผู้ใช้สามารถพิมพ์ Username นี้เพื่อเข้าสู่ระบบผู้ดูแลได้ทันที</p><button className="primary-button"><Check size={16}/> บันทึก Username ผู้ดูแล</button></form></section>
}

function SystemSettingsPanel({settings,onSaveSettings}:{settings:SystemSettings;onSaveSettings:(settings:SystemSettings)=>void}) {
  const [draft,setDraft] = useState(settings)
  useEffect(()=>setDraft(settings),[settings])
  return <section className="system-settings admin-panel"><div className="section-title"><Pencil size={18}/><div><h3>ชื่อระบบและภาคเรียน</h3><p>เปลี่ยนข้อมูลที่แสดงในระบบเมื่อต้องการ</p></div></div><div className="settings-grid"><label>ชื่อระบบ<input value={draft.title} onChange={(event)=>setDraft({...draft,title:event.target.value})}/></label><label>โรงเรียน<input className="readonly-field" value={draft.school} readOnly/></label><label>ภาคเรียน<input value={draft.term} onChange={(event)=>setDraft({...draft,term:event.target.value})}/></label><label>ปีการศึกษา<input value={draft.year} onChange={(event)=>setDraft({...draft,year:event.target.value})}/></label></div><button className="primary-button" onClick={()=>onSaveSettings({...draft,school:defaultSettings.school})}><Check size={16}/> บันทึกชื่อระบบ</button></section>
}
function ImagePickerField(){
  const picker=imagePickerController
  const [showImage,setShowImage]=useState(false)
  if(!picker)return null
  return <div className="field image-field"><span>รูปภาพคุณครูประจำชุมชน<i>*</i></span><div className="image-picker"><input id="community-image-file" type="file" accept="image/*" onChange={(e)=>picker.choose(e.target.files?.[0] ?? null)}/><input id="community-image-camera" type="file" accept="image/*" capture="environment" onChange={(e)=>picker.choose(e.target.files?.[0] ?? null)}/><div className="image-picker-actions"><label htmlFor="community-image-file"><ImagePlus size={16}/> เลือกรูป</label><label htmlFor="community-image-camera"><Camera size={16}/> ถ่ายรูป</label></div>{picker.preview&&<div className="image-preview"><button type="button" className="image-preview-trigger" onClick={()=>setShowImage(true)} aria-label="ดูรูปภาพขนาดเต็ม"><img src={picker.preview} alt="ตัวอย่างรูปคุณครูประจำชุมชน"/></button><button type="button" onClick={picker.clear}><X size={14}/> ลบรูป</button></div>}<small>{picker.preview?'คลิกรูปเพื่อดูขนาดเต็ม · ระบบจะย่อรูปไม่เกิน 500 KB และเก็บเมื่อกดบันทึกข้อมูลเท่านั้น':'ระบบจะย่อรูปไม่เกิน 500 KB และเก็บเมื่อกดบันทึกข้อมูลเท่านั้น'}</small></div>{showImage&&picker.preview&&<ImageLightbox src={picker.preview} onClose={()=>setShowImage(false)}/>}</div>
}
function Field({label,required,className='',children}:{label:string;required?:boolean;className?:string;children:React.ReactNode}){return label==='ลิงก์รูปภาพ (ไม่บังคับ)'?<ImagePickerField/>:<label className={`field ${className}`}><span>{label}{required&&<i>*</i>}</span>{children}</label>}
function Metric({value,label,icon}:{value:number|string;label:string;icon:React.ReactNode}){return <div className="metric"><span>{icon}</span><div><b>{typeof value==='number'?number(value):value}</b><small>{label}</small></div></div>}
const CommunityCard=({community,ref}:{community:Community;ref:React.Ref<HTMLDivElement>})=><article className="community-card" ref={ref}><div className="card-noise"/><div className="card-left"><div className="school-badge"><span>✦</span><small>TEACHER<br/>COMMUNITY</small></div><div className="card-title"><span>กิจกรรม</span><strong>ชุมนุม</strong><small>{community.school_name}</small></div><div className="photo-frame">{community.image_url?<img src={community.image_url} alt={community.community_name}/>:<div className="photo-placeholder"><ImagePlus size={37}/><span>ภาพกิจกรรม<br/>ของชุมชน</span></div>}</div></div><div className="card-right"><div className="card-row code"><span className="card-icon"><KeyRound size={16}/></span><div><label>รหัสชุมชน</label><b>{community.community_code}</b></div></div><div className="card-row"><span className="card-icon orange"><BookOpen size={16}/></span><div><label>ชื่อชุมชน</label><b>{community.community_name}</b></div></div><div className="card-row"><span className="card-icon pink"><Users size={16}/></span><div><label>ครูที่ปรึกษา</label><b>{community.advisor_name}</b></div></div><div className="card-split"><div><label><MapPin size={13}/> สถานที่</label><b>{community.location||'-'}</b></div><div><label><Users size={13}/> จำนวนที่รับ</label><b>{community.member_count||'-'}</b></div></div><div className="card-description"><label>รายละเอียด</label><p>{community.description||'ร่วมสร้างพื้นที่เรียนรู้ที่เปี่ยมด้วยแรงบันดาลใจ'}</p></div></div></article>
export default App
