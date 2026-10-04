import { useEffect, useMemo, useRef, useState } from 'react'
import * as htmlToImage from 'html-to-image'
import * as XLSX from 'xlsx'
import { ArrowRight, BookOpen, Check, Download, FileSpreadsheet, ImagePlus, KeyRound, LayoutDashboard, LoaderCircle, LockKeyhole, LogIn, LogOut, MapPin, Pencil, Plus, Search, Sparkles, Trash2, UserCog, Users, X } from 'lucide-react'
import { supabase, hasSupabaseConfig } from './lib/supabase'
import { seededTeachers } from './data/teachers'
import './App.css'
import './account.css'
import './safety-theme.css'

type Community = { id:string; community_code:string; community_name:string; advisor_name:string; school_name:string; location:string; member_count:number; description:string; image_url?:string; owner_id?:string }
type TeacherAccount = { id:string; teacher_name:string; community_code:string; role:'teacher'|'admin'; community_id?:string; login_email?:string }
const emptyForm = { community_name:'', advisor_name:'', school_name:'', location:'', member_count:'', description:'', image_url:'' }
const emptyAccount = { teacher_name:'', community_code:'' }
const demos: Community[] = [
  { id:'demo-1', community_code:'กก026', community_name:'ร้านค้าสวัสดิการโรงเรียน', advisor_name:'นางจุฑาวรรณ ลิทธิชัย', school_name:'โรงเรียนนิซัยมาตู', location:'ร้านค้าสวัสดิการฯ', member_count:20, description:'เรียนรู้เกี่ยวกับการจัดการร้านค้าขายปลีก ฝึกการทำงานเป็นทีมและความรับผิดชอบ', owner_id:'teacher-1' },
  { id:'demo-2', community_code:'กก027', community_name:'นักสร้างสื่อสร้างสรรค์', advisor_name:'นายธนกร วัฒนชัย', school_name:'โรงเรียนนิซัยมาตู', location:'ห้องคอมพิวเตอร์ 2', member_count:16, description:'ฝึกออกแบบสื่อการเรียนรู้และสื่อประชาสัมพันธ์อย่างสร้างสรรค์', owner_id:'teacher-2' },
]
const demoAccounts: TeacherAccount[] = [{ id:'admin-1', teacher_name:'ผู้ดูแลระบบ', community_code:'admin2569', role:'admin' }, ...seededTeachers]
type SystemSettings = { title:string; term:string; year:string; school:string; adminPassword:string }
const defaultSettings:SystemSettings = { title:'ระบบลงทะเบียนชุมนุมคุณครู', term:'ภาคเรียนที่ 2', year:'ปีการศึกษา 2569', school:'โรงเรียนวิเชียรมาตุ', adminPassword:'admin2569' }
const number = (n:number) => new Intl.NumberFormat('th-TH').format(n)

function App() {
  const [communities,setCommunities] = useState<Community[]>([])
  const [accounts,setAccounts] = useState<TeacherAccount[]>([])
  const [current,setCurrent] = useState<TeacherAccount | null>(null)
  const [view,setView] = useState<'teacher'|'admin'>('teacher')
  const [form,setForm] = useState(emptyForm)
  const [editingId,setEditingId] = useState<string|null>(null)
  const [selectedId,setSelectedId] = useState<string|null>(null)
  const [accountForm,setAccountForm] = useState(emptyAccount)
  const [editingAccount,setEditingAccount] = useState<string|null>(null)
  const [teacherSearch,setTeacherSearch] = useState('')
  const [adminPassword,setAdminPassword] = useState('')
  const [search,setSearch] = useState('')
  const [showLogin,setShowLogin] = useState(false)
  const [showAdminLogin,setShowAdminLogin] = useState(false)
  const [settings,setSettings] = useState<SystemSettings>(defaultSettings)
  const [notice,setNotice] = useState('')
  const [loading,setLoading] = useState(true)
  const cardRef = useRef<HTMLDivElement>(null)
  const isAdmin = current?.role === 'admin'
  const selected = communities.find((item) => item.id === selectedId) ?? communities[0]
  const filtered = useMemo(() => communities.filter((item) => [item.community_code,item.community_name,item.advisor_name].join(' ').toLowerCase().includes(search.toLowerCase())),[communities,search])
  const teacherMatches = useMemo(() => accounts.filter((item) => item.role==='teacher' && item.teacher_name.includes(teacherSearch.trim())).slice(0,9),[accounts,teacherSearch])
  useEffect(() => { void load() }, [])
  const flash = (text:string) => { setNotice(text); window.setTimeout(() => setNotice(''),4500) }
  async function load() {
    setLoading(true)
    if (!hasSupabaseConfig) {
      const localCommunities = JSON.parse(localStorage.getItem('teacher-community-demo') ?? 'null') as Community[] | null
      const localAccounts = JSON.parse(localStorage.getItem('teacher-community-accounts') ?? 'null') as TeacherAccount[] | null
      const migratedAccounts = !localAccounts || localAccounts.filter((item)=>item.role==='teacher').length < 100 ? demoAccounts : localAccounts
      const localSettings = JSON.parse(localStorage.getItem('teacher-community-settings') ?? 'null') as SystemSettings | null
      setCommunities(localCommunities ?? demos); setAccounts(migratedAccounts); setSettings(localSettings ?? defaultSettings); setSelectedId((localCommunities ?? demos)[0].id); setLoading(false); return
    }
    const [{data: dataCommunities,error:communityError},{data:profile}] = await Promise.all([
      supabase!.from('communities').select('*').order('created_at'), supabase!.auth.getUser(),
    ])
    if (communityError) flash(`โหลดข้อมูลไม่สำเร็จ: ${communityError.message}`)
    setCommunities((dataCommunities ?? []) as Community[]); setSelectedId(dataCommunities?.[0]?.id ?? null)
    if (profile.user) {
      const {data: account} = await supabase!.from('teacher_profiles').select('*').eq('id',profile.user.id).single()
      setCurrent(account as TeacherAccount | null)
      if (account?.role === 'admin') { const {data} = await supabase!.from('teacher_profiles').select('*').order('teacher_name'); setAccounts((data ?? []) as TeacherAccount[]) }
    }
    setLoading(false)
  }
  function setField(key:keyof typeof emptyForm,value:string) { setForm((now) => ({...now,[key]:value})) }
  function openEditor(community?:Community) {
    if (!current && hasSupabaseConfig) { setShowLogin(true); flash('กรุณาเข้าสู่ระบบก่อนแก้ไขข้อมูล'); return }
    if (community) { setEditingId(community.id); setForm({community_name:community.community_name,advisor_name:community.advisor_name,school_name:community.school_name,location:community.location,member_count:String(community.member_count),description:community.description,image_url:community.image_url ?? ''}) } else { setEditingId(null); setForm({...emptyForm,advisor_name:current?.teacher_name ?? ''}) }
    setView('teacher'); window.scrollTo({top:0,behavior:'smooth'})
  }
  async function saveCommunity(event:React.FormEvent) {
    event.preventDefault()
    if (!current) { setShowLogin(true); flash('กรุณาเข้าสู่ระบบก่อนบันทึก'); return }
    if (!form.community_name || !form.advisor_name || !form.school_name) { flash('กรุณากรอกชื่อชุมชน ครูที่ปรึกษา และโรงเรียน'); return }
    const payload = {community_name:form.community_name.trim(),advisor_name:form.advisor_name.trim(),school_name:form.school_name.trim(),location:form.location.trim(),member_count:Number(form.member_count)||0,description:form.description.trim(),image_url:form.image_url.trim()||null}
    if (!hasSupabaseConfig) {
      const code = current.community_code
      const item:Community = editingId ? {...communities.find((x)=>x.id===editingId)!, ...payload, image_url:payload.image_url ?? undefined} : {...payload,id:crypto.randomUUID(),community_code:code,owner_id:current.id,image_url:payload.image_url ?? undefined}
      const next = editingId ? communities.map((x)=>x.id===item.id?item:x) : [...communities,item]
      setCommunities(next); localStorage.setItem('teacher-community-demo',JSON.stringify(next));
      if (!editingId) { const nextAccounts=accounts.map((x)=>x.id===current.id?{...x,community_id:item.id}:x); setAccounts(nextAccounts); setCurrent(nextAccounts.find((x)=>x.id===current.id) ?? current); localStorage.setItem('teacher-community-accounts',JSON.stringify(nextAccounts)) }
      setSelectedId(item.id); setEditingId(null); setForm(emptyForm); flash('บันทึกข้อมูลชุมชนแล้ว'); return
    }
    const answer = editingId ? await supabase!.from('communities').update(payload).eq('id',editingId).select().single() : await supabase!.from('communities').insert({...payload,community_code:current.community_code,owner_id:current.id}).select().single()
    if (answer.error) { flash(`บันทึกไม่สำเร็จ: ${answer.error.message}`); return }
    await load(); setSelectedId(answer.data.id); setEditingId(null); setForm(emptyForm); flash('บันทึกข้อมูลชุมชนแล้ว')
  }
  function selectTeacher(account:TeacherAccount) { setCurrent(account); setShowLogin(false); setTeacherSearch(''); flash(`ยินดีต้อนรับ ${account.teacher_name}`) }
  function adminSignIn(event:React.FormEvent) { event.preventDefault(); if (adminPassword !== settings.adminPassword) { flash('รหัสผู้ดูแลไม่ถูกต้อง'); return }; setCurrent(accounts.find((item)=>item.role==='admin') ?? {id:'admin-1',teacher_name:'ผู้ดูแลระบบ',community_code:'',role:'admin'}); setShowAdminLogin(false); setAdminPassword(''); setView('admin'); flash('เข้าสู่ระบบผู้ดูแลแล้ว') }
  function updateSettings(next:SystemSettings) { setSettings(next); localStorage.setItem('teacher-community-settings',JSON.stringify(next)); flash('บันทึกชื่อระบบแล้ว') }
  function signOut() { if (hasSupabaseConfig) void supabase!.auth.signOut(); setCurrent(null); setView('teacher'); flash('ออกจากระบบแล้ว') }
  async function saveAccount(event:React.FormEvent) {
    event.preventDefault(); if (!accountForm.teacher_name || !accountForm.community_code) { flash('กรุณากรอกชื่อครูและรหัสชุมชน'); return }
    if (!hasSupabaseConfig) {
      const duplicate=accounts.some((x)=>x.id!==editingAccount&&(x.teacher_name===accountForm.teacher_name.trim()||x.community_code.toLowerCase()===accountForm.community_code.trim().toLowerCase()))
      if (duplicate) { flash('ชื่อครูหรือรหัสชุมชนนี้ถูกใช้แล้ว'); return }
      const item:TeacherAccount = editingAccount ? {...accounts.find((x)=>x.id===editingAccount)!, ...accountForm} : {id:crypto.randomUUID(),...accountForm,role:'teacher'}
      const next=editingAccount?accounts.map((x)=>x.id===item.id?item:x):[...accounts,item]; setAccounts(next); localStorage.setItem('teacher-community-accounts',JSON.stringify(next)); setAccountForm(emptyAccount); setEditingAccount(null); flash(editingAccount?'แก้ไขบัญชีครูแล้ว':'เพิ่มบัญชีครูแล้ว'); return
    }
    const {error} = await supabase!.functions.invoke('admin-accounts',{body:{action:editingAccount?'update':'create',id:editingAccount,teacher_name:accountForm.teacher_name.trim(),community_code:accountForm.community_code.trim()}})
    if (error) { flash(`บันทึกไม่สำเร็จ: ${error.message}`); return }; setAccountForm(emptyAccount); setEditingAccount(null); await load(); flash('บันทึกบัญชีครูแล้ว')
  }
  async function deleteAccount(account:TeacherAccount) {
    if (!window.confirm(`ลบบัญชี ${account.teacher_name} ใช่หรือไม่?`)) return
    if (!hasSupabaseConfig) { const next=accounts.filter((x)=>x.id!==account.id); setAccounts(next); localStorage.setItem('teacher-community-accounts',JSON.stringify(next)); flash('ลบบัญชีครูแล้ว'); return }
    const {error}=await supabase!.functions.invoke('admin-accounts',{body:{action:'delete',id:account.id}}); if(error){flash(`ลบไม่สำเร็จ: ${error.message}`);return}; await load();flash('ลบบัญชีครูแล้ว')
  }
  function editAccount(account:TeacherAccount) { setEditingAccount(account.id); setAccountForm({teacher_name:account.teacher_name,community_code:account.community_code}) }
  function exportExcel() { const sheet=XLSX.utils.json_to_sheet(communities.map((x)=>({'รหัสชุมชน':x.community_code,'ชื่อชุมชน':x.community_name,'ครูที่ปรึกษา':x.advisor_name,'โรงเรียน':x.school_name,'สถานที่':x.location,'จำนวนสมาชิก':x.member_count,'รายละเอียด':x.description}))); sheet['!cols']=[12,30,24,26,20,14,60].map((wch)=>({wch})); const book=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book,sheet,'ชุมชนคุณครู'); XLSX.writeFile(book,'teacher-community.xlsx') }
  async function exportCard() { if(!cardRef.current||!selected)return; const href=await htmlToImage.toPng(cardRef.current,{pixelRatio:2,backgroundColor:'#fff7eb'});const a=document.createElement('a');a.href=href;a.download=`${selected.community_code}-${selected.community_name}.png`;a.click() }
  return <main className="app-shell">
    <header className="topbar safety-topbar"><button className="brand" onClick={()=>setView('teacher')}><span className="brand-mark"><Sparkles size={19}/></span><span>{settings.title}<small>{settings.term} · {settings.year}</small></span></button><nav className="top-actions"><button className={view==='teacher'?'nav-button active':'nav-button'} onClick={()=>setView('teacher')}><Pencil size={16}/> ข้อมูลชุมชน</button><button className={view==='admin'?'nav-button active':'nav-button'} onClick={()=>isAdmin?setView('admin'):setShowAdminLogin(true)}><LayoutDashboard size={16}/> ผู้ดูแลระบบ</button>{current?<button className="login-button" onClick={signOut}><LogOut size={16}/> ออกจากระบบ</button>:<button className="login-button" onClick={()=>setShowLogin(true)}><LogIn size={16}/> ค้นหาชื่อครู</button>}</nav></header>
    {notice&&<div className="notice"><Check size={17}/>{notice}<button onClick={()=>setNotice('')}><X size={16}/></button></div>}
    {view==='teacher'?<section className="teacher-page safety-hero"><div className="teacher-intro"><div className="eyebrow"><Sparkles size={15}/> {settings.school}</div><h1>{current?<>สวัสดี<br/><em>{current.teacher_name}</em></>:<>{settings.title}<br/><em>{settings.term}</em></>}</h1><p>{current?`รหัสชุมนุมของคุณคือ ${current.community_code} คุณสามารถบันทึกหรือแก้ไขข้อมูลชุมนุมของตนเองได้`:'พิมพ์ชื่อของคุณเพื่อค้นหารายชื่อ แล้วกดเลือกเพื่อเข้าใช้งานได้ทันที'}</p><div className="intro-stats"><span><Users size={17}/>{number(accounts.filter((item)=>item.role==='teacher').length)} รายชื่อครู</span><span><KeyRound size={17}/> รหัสชุมนุมแสดงก่อนเข้า</span></div>{current&&communities.find((x)=>x.owner_id===current.id)&&<button className="secondary-action" onClick={()=>openEditor(communities.find((x)=>x.owner_id===current.id))}><Pencil size={15}/> แก้ไขชุมนุมของฉัน</button>}</div><form className="community-form" onSubmit={saveCommunity}><div className="form-heading"><div className="form-icon"><BookOpen size={21}/></div><div><h2>{editingId?'แก้ไขข้อมูลชุมนุม':'ข้อมูลชุมนุมของคุณ'}</h2><p>{current?`บันทึกภายใต้รายชื่อ ${current.teacher_name}`:'ค้นหาชื่อครูของคุณก่อนจึงจะบันทึกข้อมูลได้'}</p></div></div><fieldset disabled={!current} className="form-grid"><Field label="ชื่อชุมนุม" required><input value={form.community_name} onChange={(e)=>setField('community_name',e.target.value)} placeholder="เช่น ร้านค้าสวัสดิการโรงเรียน"/></Field><Field label="ครูที่ปรึกษา" required><input value={form.advisor_name} onChange={(e)=>setField('advisor_name',e.target.value)} placeholder="ชื่อ-นามสกุล"/></Field><Field label="โรงเรียน" required><input value={form.school_name} onChange={(e)=>setField('school_name',e.target.value)} placeholder={settings.school}/></Field><Field label="สถานที่ดำเนินกิจกรรม"><input value={form.location} onChange={(e)=>setField('location',e.target.value)} placeholder="เช่น ห้องคอมพิวเตอร์ 2"/></Field><Field label="จำนวนสมาชิก"><input type="number" min="0" value={form.member_count} onChange={(e)=>setField('member_count',e.target.value)} placeholder="0"/></Field><Field label="ลิงก์รูปภาพ (ไม่บังคับ)"><input value={form.image_url} onChange={(e)=>setField('image_url',e.target.value)} placeholder="https://..."/></Field><Field label="รายละเอียดกิจกรรม" className="span-2"><textarea rows={4} value={form.description} onChange={(e)=>setField('description',e.target.value)} placeholder="อธิบายเป้าหมาย กิจกรรม หรือสิ่งที่นักเรียนได้เรียนรู้"/></Field></fieldset><div className="form-footer"><span><LockKeyhole size={15}/>{current?'ข้อมูลนี้แก้ไขได้เฉพาะรายชื่อของคุณ':'ข้อมูลจะถูกปลดล็อกเมื่อเลือกรายชื่อแล้ว'}</span><div>{current?<><button type="button" className="text-button" onClick={()=>{setForm(emptyForm);setEditingId(null)}}>ล้างข้อมูล</button><button className="primary-button"><ArrowRight size={17}/>{editingId?'บันทึกการแก้ไข':'บันทึกชุมนุม'}</button></>:<button type="button" className="primary-button" onClick={()=>setShowLogin(true)}><Search size={17}/> ค้นหาชื่อครู</button>}</div></div></form></section>:<AdminPage accounts={accounts} communities={communities} filtered={filtered} selected={selected} search={search} setSearch={setSearch} loading={loading} accountForm={accountForm} setAccountForm={setAccountForm} editingAccount={editingAccount} onSaveAccount={saveAccount} onEditAccount={editAccount} onDeleteAccount={deleteAccount} onCancelAccount={()=>{setEditingAccount(null);setAccountForm(emptyAccount)}} onExportExcel={exportExcel} onExportCard={exportCard} onSelected={setSelectedId} cardRef={cardRef} settings={settings} onSaveSettings={updateSettings}/>} 
    {showLogin&&<div className="modal-backdrop" onMouseDown={()=>setShowLogin(false)}><section className="login-modal teacher-picker" onMouseDown={(e)=>e.stopPropagation()}><button className="modal-close" onClick={()=>setShowLogin(false)}><X size={19}/></button><div className="login-symbol"><Search size={25}/></div><h2>ค้นหารายชื่อครู</h2><p>พิมพ์ชื่อ-นามสกุล แล้วเลือกรายชื่อของคุณเพื่อเข้าใช้งาน</p><label>ชื่อครู<input autoFocus value={teacherSearch} onChange={(e)=>setTeacherSearch(e.target.value)} placeholder="เริ่มพิมพ์ชื่อครู..."/></label><div className="teacher-results">{teacherSearch.trim()?teacherMatches.map((teacher)=><button key={teacher.id} onClick={()=>selectTeacher(teacher)}><span>{teacher.teacher_name.slice(0,1)}</span><div><b>{teacher.teacher_name}</b><small><KeyRound size={12}/> รหัสชุมนุม {teacher.community_code}</small></div><ArrowRight size={17}/></button>):<div className="search-empty">พิมพ์ชื่อเพื่อค้นหารายชื่อครู {number(accounts.filter((item)=>item.role==='teacher').length)} รายชื่อ</div>}{teacherSearch.trim()&&teacherMatches.length===0&&<div className="search-empty">ไม่พบรายชื่อ ลองพิมพ์คำอื่น หรือแจ้งผู้ดูแลระบบ</div>}</div></section></div>}
    {showAdminLogin&&<div className="modal-backdrop" onMouseDown={()=>setShowAdminLogin(false)}><section className="login-modal admin-login" onMouseDown={(e)=>e.stopPropagation()}><button className="modal-close" onClick={()=>setShowAdminLogin(false)}><X size={19}/></button><div className="login-symbol"><LockKeyhole size={25}/></div><h2>เข้าสู่ระบบผู้ดูแล</h2><p>สำหรับจัดการรายชื่อครู รหัสชุมนุม และชื่อระบบ</p><form onSubmit={adminSignIn}><label>รหัสผู้ดูแล<input autoFocus type="password" value={adminPassword} onChange={(e)=>setAdminPassword(e.target.value)} placeholder="กรอกรหัสผู้ดูแล" required/></label><button className="primary-button full"><LockKeyhole size={17}/> เข้าสู่ระบบผู้ดูแล</button></form></section></div>}
  </main>
}

type AdminProps={accounts:TeacherAccount[];communities:Community[];filtered:Community[];selected?:Community;search:string;setSearch:(x:string)=>void;loading:boolean;accountForm:typeof emptyAccount;setAccountForm:(x:typeof emptyAccount)=>void;editingAccount:string|null;onSaveAccount:(e:React.FormEvent)=>void;onEditAccount:(x:TeacherAccount)=>void;onDeleteAccount:(x:TeacherAccount)=>void;onCancelAccount:()=>void;onExportExcel:()=>void;onExportCard:()=>void;onSelected:(x:string)=>void;cardRef:React.Ref<HTMLDivElement>;settings:SystemSettings;onSaveSettings:(x:SystemSettings)=>void}
function AdminPage(p:AdminProps){const [draft,setDraft]=useState(p.settings);useEffect(()=>setDraft(p.settings),[p.settings]);return <section className="admin-page safety-admin"><div className="admin-header"><div><div className="eyebrow"><LayoutDashboard size={15}/> ศูนย์ควบคุมผู้ดูแล</div><h1>{p.settings.title}</h1><p>{p.settings.school} · {p.settings.term} · {p.settings.year}</p></div><button className="export-button" onClick={p.onExportExcel}><FileSpreadsheet size={17}/> Export Excel</button></div><div className="admin-metrics"><Metric value={p.communities.length} label="ชุมนุมที่ลงทะเบียน" icon={<BookOpen/>}/><Metric value={p.accounts.filter((x)=>x.role==='teacher').length} label="รายชื่อครู" icon={<Users/>}/><Metric value="กก001–125" label="รหัสชุมนุมจากไฟล์" icon={<KeyRound/>}/></div><div className="admin-columns"><section className="account-manager"><div className="section-title"><UserCog size={18}/><div><h3>จัดการรายชื่อครู</h3><p>ครูค้นหาและเลือกชื่อตัวเองได้ทันที</p></div></div><form className="account-form" onSubmit={p.onSaveAccount}><label>ชื่อครู<input value={p.accountForm.teacher_name} onChange={(e)=>p.setAccountForm({...p.accountForm,teacher_name:e.target.value})} placeholder="ชื่อ-นามสกุลครู"/></label><label>รหัสชุมนุม<input value={p.accountForm.community_code} onChange={(e)=>p.setAccountForm({...p.accountForm,community_code:e.target.value})} placeholder="เช่น กก126"/></label><div><button className="primary-button"><Plus size={16}/>{p.editingAccount?'บันทึกรายชื่อ':'เพิ่มรายชื่อ'}</button>{p.editingAccount&&<button type="button" className="text-button" onClick={p.onCancelAccount}>ยกเลิก</button>}</div></form><div className="account-list">{p.accounts.filter((x)=>x.role==='teacher').map((a)=><div className="account-row" key={a.id}><span className="account-avatar">{a.teacher_name.slice(0,1)}</span><div><b>{a.teacher_name}</b><small><KeyRound size={11}/>{a.community_code}</small></div><button onClick={()=>p.onEditAccount(a)} title="แก้ไข"><Pencil size={15}/></button><button className="danger" onClick={()=>p.onDeleteAccount(a)} title="ลบ"><Trash2 size={15}/></button></div>)}</div></section><div><section className="system-settings"><div className="section-title"><Pencil size={18}/><div><h3>ชื่อระบบและภาคเรียน</h3><p>เปลี่ยนได้ทุกครั้งที่เปิดภาคเรียนใหม่</p></div></div><div className="settings-grid"><label>ชื่อระบบ<input value={draft.title} onChange={(e)=>setDraft({...draft,title:e.target.value})}/></label><label>โรงเรียน<input value={draft.school} onChange={(e)=>setDraft({...draft,school:e.target.value})}/></label><label>ภาคเรียน<input value={draft.term} onChange={(e)=>setDraft({...draft,term:e.target.value})}/></label><label>ปีการศึกษา<input value={draft.year} onChange={(e)=>setDraft({...draft,year:e.target.value})}/></label><label className="span-2">รหัสผู้ดูแล<input type="password" value={draft.adminPassword} onChange={(e)=>setDraft({...draft,adminPassword:e.target.value})}/></label></div><button className="primary-button" onClick={()=>p.onSaveSettings(draft)}><Check size={16}/> บันทึกชื่อระบบ</button></section><section className="community-preview"><div className="list-head"><h3>ตัวอย่างภาพชุมนุม</h3><button className="export-image" onClick={p.onExportCard}><Download size={16}/> PNG</button></div><label className="search"><Search size={17}/><input value={p.search} onChange={(e)=>p.setSearch(e.target.value)} placeholder="ค้นหาชุมนุม..."/></label><div className="admin-workspace"><aside className="community-list"><div className="list-items">{p.loading?<div className="loading"><LoaderCircle className="spin"/>กำลังโหลด...</div>:p.filtered.map((c)=><button key={c.id} className={c.id===p.selected?.id?'community-item selected':'community-item'} onClick={()=>p.onSelected(c.id)}><span className="item-avatar">{c.community_name.slice(0,1)}</span><span><b>{c.community_name}</b><small>{c.community_code} · {c.member_count} สมาชิก</small></span></button>)}</div></aside><div className="preview-pane">{p.selected?<CommunityCard community={p.selected} ref={p.cardRef}/>:<div className="empty">ยังไม่มีข้อมูลชุมนุม</div>}</div></div></section></div></div></section>}
function Field({label,required,className='',children}:{label:string;required?:boolean;className?:string;children:React.ReactNode}){return <label className={`field ${className}`}><span>{label}{required&&<i>*</i>}</span>{children}</label>}
function Metric({value,label,icon}:{value:number|string;label:string;icon:React.ReactNode}){return <div className="metric"><span>{icon}</span><div><b>{typeof value==='number'?number(value):value}</b><small>{label}</small></div></div>}
const CommunityCard=({community,ref}:{community:Community;ref:React.Ref<HTMLDivElement>})=><article className="community-card" ref={ref}><div className="card-noise"/><div className="card-left"><div className="school-badge"><span>✦</span><small>TEACHER<br/>COMMUNITY</small></div><div className="card-title"><span>กิจกรรม</span><strong>ชุมนุม</strong><small>{community.school_name}</small></div><div className="photo-frame">{community.image_url?<img src={community.image_url} alt={community.community_name}/>:<div className="photo-placeholder"><ImagePlus size={37}/><span>ภาพกิจกรรม<br/>ของชุมชน</span></div>}</div></div><div className="card-right"><div className="card-row code"><span className="card-icon"><KeyRound size={16}/></span><div><label>รหัสชุมชน</label><b>{community.community_code}</b></div></div><div className="card-row"><span className="card-icon orange"><BookOpen size={16}/></span><div><label>ชื่อชุมชน</label><b>{community.community_name}</b></div></div><div className="card-row"><span className="card-icon pink"><Users size={16}/></span><div><label>ครูที่ปรึกษา</label><b>{community.advisor_name}</b></div></div><div className="card-split"><div><label><MapPin size={13}/> สถานที่</label><b>{community.location||'-'}</b></div><div><label><Users size={13}/> จำนวนที่รับ</label><b>{community.member_count||'-'}</b></div></div><div className="card-description"><label>รายละเอียด</label><p>{community.description||'ร่วมสร้างพื้นที่เรียนรู้ที่เปี่ยมด้วยแรงบันดาลใจ'}</p></div></div></article>
export default App
