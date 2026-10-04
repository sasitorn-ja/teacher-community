import { useEffect, useMemo, useRef, useState } from 'react'
import * as htmlToImage from 'html-to-image'
import * as XLSX from 'xlsx'
import {
  ArrowRight, BookOpen, Check, Download, Edit3, FileSpreadsheet, ImagePlus,
  KeyRound, LayoutDashboard, LoaderCircle, LockKeyhole, LogIn,
  MapPin, Pencil, Plus, Search, Sparkles, Users, X,
} from 'lucide-react'
import { supabase, hasSupabaseConfig } from './lib/supabase'
import './App.css'

type Community = {
  id: string
  community_code: string
  community_name: string
  advisor_name: string
  school_name: string
  location: string
  member_count: number
  description: string
  image_url?: string
}

const emptyForm = { community_name: '', advisor_name: '', school_name: '', location: '', member_count: '', description: '', image_url: '' }

const demoCommunities: Community[] = [
  { id: 'demo-1', community_code: 'กก026', community_name: 'ร้านค้าสวัสดิการโรงเรียน', advisor_name: 'นางจุฑาวรรณ ลิทธิชัย', school_name: 'โรงเรียนนิซัยมาตู', location: 'ร้านค้าสวัสดิการฯ', member_count: 20, description: 'เรียนรู้เกี่ยวกับการจัดการร้านค้าขายปลีก ฝึกการทำงานเป็นทีมและความรับผิดชอบ' },
  { id: 'demo-2', community_code: 'กก027', community_name: 'นักสร้างสื่อสร้างสรรค์', advisor_name: 'นายธนกร วัฒนชัย', school_name: 'โรงเรียนนิซัยมาตู', location: 'ห้องคอมพิวเตอร์ 2', member_count: 16, description: 'ฝึกออกแบบสื่อการเรียนรู้และสื่อประชาสัมพันธ์อย่างสร้างสรรค์' },
]

const thaiNumber = (value: number) => new Intl.NumberFormat('th-TH').format(value)

function App() {
  const [communities, setCommunities] = useState<Community[]>([])
  const [view, setView] = useState<'teacher' | 'admin'>('teacher')
  const [form, setForm] = useState(emptyForm)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [search, setSearch] = useState('')
  const [isLoading, setIsLoading] = useState(true)
  const [isSaving, setIsSaving] = useState(false)
  const [notice, setNotice] = useState('')
  const [showLogin, setShowLogin] = useState(false)
  const [email, setEmail] = useState('')
  const [loginMessage, setLoginMessage] = useState('')
  const cardRef = useRef<HTMLDivElement>(null)
  const selected = communities.find((community) => community.id === selectedId) ?? communities[0]
  const filtered = useMemo(() => communities.filter((community) => [community.community_code, community.community_name, community.advisor_name, community.school_name].join(' ').toLowerCase().includes(search.toLowerCase())), [communities, search])

  useEffect(() => { void loadCommunities() }, [])

  async function loadCommunities() {
    setIsLoading(true)
    if (!hasSupabaseConfig) {
      const stored = localStorage.getItem('teacher-community-demo')
      const local = stored ? JSON.parse(stored) as Community[] : demoCommunities
      setCommunities(local); setSelectedId(local[0]?.id ?? null); setIsLoading(false); return
    }
    const { data, error } = await supabase!.from('communities').select('*').order('created_at', { ascending: true })
    if (error) setNotice(`ไม่สามารถโหลดข้อมูลได้: ${error.message}`)
    const list = (data ?? []) as Community[]
    setCommunities(list); setSelectedId(list[0]?.id ?? null); setIsLoading(false)
  }
  function updateForm(key: keyof typeof emptyForm, value: string) { setForm((current) => ({ ...current, [key]: value })) }
  function beginEdit(community?: Community) {
    if (community) { setEditingId(community.id); setForm({ community_name: community.community_name, advisor_name: community.advisor_name, school_name: community.school_name, location: community.location, member_count: String(community.member_count), description: community.description, image_url: community.image_url ?? '' }) } else { setEditingId(null); setForm(emptyForm) }
    setView('teacher'); window.scrollTo({ top: 0, behavior: 'smooth' })
  }
  async function saveCommunity(event: React.FormEvent) {
    event.preventDefault()
    if (!form.community_name || !form.advisor_name || !form.school_name) { setNotice('กรุณากรอกชื่อชุมชน ครูที่ปรึกษา และชื่อโรงเรียน'); return }
    setIsSaving(true)
    const payload = { community_name: form.community_name.trim(), advisor_name: form.advisor_name.trim(), school_name: form.school_name.trim(), location: form.location.trim(), member_count: Number(form.member_count) || 0, description: form.description.trim(), image_url: form.image_url.trim() || null }
    if (!hasSupabaseConfig) {
      const existing = communities.find((community) => community.id === editingId)
      const newCommunity: Community = existing ? { ...existing, ...payload, image_url: payload.image_url ?? undefined } : { ...payload, image_url: payload.image_url ?? undefined, id: crypto.randomUUID(), community_code: `กก${String(26 + communities.length).padStart(3, '0')}` }
      const updated = existing ? communities.map((community) => community.id === editingId ? newCommunity : community) : [...communities, newCommunity]
      localStorage.setItem('teacher-community-demo', JSON.stringify(updated)); setCommunities(updated); setSelectedId(newCommunity.id); completeSave(existing ? 'แก้ไขข้อมูลชุมชนแล้ว' : `ลงทะเบียนสำเร็จ — รหัสชุมชน ${newCommunity.community_code}`); return
    }
    const { data: authData } = await supabase!.auth.getUser()
    if (!authData.user) { setNotice('กรุณาเข้าสู่ระบบก่อนลงทะเบียนหรือแก้ไขข้อมูล'); setShowLogin(true); setIsSaving(false); return }
    const response = editingId ? await supabase!.from('communities').update(payload).eq('id', editingId).select().single() : await supabase!.from('communities').insert({ ...payload, owner_id: authData.user.id }).select().single()
    if (response.error) { setNotice(`บันทึกไม่สำเร็จ: ${response.error.message}`); setIsSaving(false); return }
    const saved = response.data as Community
    setCommunities((current) => editingId ? current.map((item) => item.id === saved.id ? saved : item) : [...current, saved]); setSelectedId(saved.id); completeSave(editingId ? 'แก้ไขข้อมูลชุมชนแล้ว' : `ลงทะเบียนสำเร็จ — รหัสชุมชน ${saved.community_code}`)
  }
  function completeSave(message: string) { setForm(emptyForm); setEditingId(null); setIsSaving(false); setNotice(message); window.setTimeout(() => setNotice(''), 5000) }
  async function sendMagicLink(event: React.FormEvent) { event.preventDefault(); if (!hasSupabaseConfig) { setLoginMessage('โหมดตัวอย่าง: เพิ่มค่า Supabase ในไฟล์ .env เพื่อเปิดใช้การเข้าสู่ระบบจริง'); return }; const { error } = await supabase!.auth.signInWithOtp({ email, options: { emailRedirectTo: window.location.origin } }); setLoginMessage(error ? error.message : 'ส่งลิงก์เข้าสู่ระบบไปยังอีเมลแล้ว') }
  function exportExcel() { const rows = communities.map((community) => ({ 'รหัสชุมชน': community.community_code, 'ชื่อชุมชน': community.community_name, 'ครูที่ปรึกษา': community.advisor_name, 'โรงเรียน': community.school_name, 'สถานที่': community.location, 'จำนวนสมาชิก': community.member_count, 'รายละเอียด': community.description })); const worksheet = XLSX.utils.json_to_sheet(rows); worksheet['!cols'] = [12, 30, 25, 26, 25, 15, 60].map((wch) => ({ wch })); const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, worksheet, 'ชุมชนคุณครู'); XLSX.writeFile(workbook, 'teacher-community.xlsx') }
  async function exportCard() { if (!cardRef.current || !selected) return; const dataUrl = await htmlToImage.toPng(cardRef.current, { pixelRatio: 2, backgroundColor: '#fff7eb' }); const link = document.createElement('a'); link.download = `${selected.community_code}-${selected.community_name}.png`; link.href = dataUrl; link.click() }

  return <main className="app-shell">
    <header className="topbar"><button className="brand" onClick={() => setView('teacher')} aria-label="หน้าแรก"><span className="brand-mark"><Sparkles size={19} /></span><span>ชุมชนคุณครู <small>TEACHER COMMUNITY</small></span></button><nav className="top-actions"><button className={view === 'teacher' ? 'nav-button active' : 'nav-button'} onClick={() => setView('teacher')}><Pencil size={16} /> ลงทะเบียน</button><button className={view === 'admin' ? 'nav-button active' : 'nav-button'} onClick={() => setView('admin')}><LayoutDashboard size={16} /> สำหรับแอดมิน</button><button className="login-button" onClick={() => setShowLogin(true)}><LogIn size={16} /> เข้าสู่ระบบ</button></nav></header>
    {notice && <div className="notice"><Check size={17} /> {notice}<button onClick={() => setNotice('')}><X size={16} /></button></div>}
    {view === 'teacher' ? <section className="teacher-page"><div className="teacher-intro"><div className="eyebrow"><Sparkles size={15} /> เปิดพื้นที่การเรียนรู้ร่วมกัน</div><h1>ลงทะเบียน<br /><em>ชุมชนคุณครู</em></h1><p>บอกเล่าเรื่องราวของชุมชนการเรียนรู้ในโรงเรียนของคุณ เพื่อให้ทุกความตั้งใจได้ถูกมองเห็น</p><div className="intro-stats"><span><Users size={17} /> {thaiNumber(communities.length)} ชุมชน</span><span><KeyRound size={17} /> รหัสเริ่ม กก026</span></div></div><form className="community-form" onSubmit={saveCommunity}><div className="form-heading"><div className="form-icon"><BookOpen size={21} /></div><div><h2>{editingId ? 'แก้ไขข้อมูลชุมชน' : 'ข้อมูลชุมชนของคุณ'}</h2><p>ช่องที่มี * จำเป็นต้องกรอก</p></div></div><div className="form-grid"><Field label="ชื่อชุมชน" required><input value={form.community_name} onChange={(event) => updateForm('community_name', event.target.value)} placeholder="เช่น ร้านค้าสวัสดิการโรงเรียน" /></Field><Field label="ครูที่ปรึกษา" required><input value={form.advisor_name} onChange={(event) => updateForm('advisor_name', event.target.value)} placeholder="ชื่อ-นามสกุล" /></Field><Field label="โรงเรียน" required><input value={form.school_name} onChange={(event) => updateForm('school_name', event.target.value)} placeholder="ชื่อโรงเรียน" /></Field><Field label="สถานที่ดำเนินกิจกรรม"><input value={form.location} onChange={(event) => updateForm('location', event.target.value)} placeholder="เช่น ห้องคอมพิวเตอร์ 2" /></Field><Field label="จำนวนสมาชิก"><input inputMode="numeric" type="number" min="0" value={form.member_count} onChange={(event) => updateForm('member_count', event.target.value)} placeholder="0" /></Field><Field label="ลิงก์รูปภาพ (ไม่บังคับ)"><input value={form.image_url} onChange={(event) => updateForm('image_url', event.target.value)} placeholder="https://..." /></Field><Field label="รายละเอียดกิจกรรม" className="span-2"><textarea rows={4} value={form.description} onChange={(event) => updateForm('description', event.target.value)} placeholder="อธิบายเป้าหมาย กิจกรรม หรือสิ่งที่นักเรียนได้เรียนรู้" /></Field></div><div className="form-footer"><span><LockKeyhole size={15} /> คุณจะแก้ไขข้อมูลของตนเองได้หลังเข้าสู่ระบบ</span><div><button type="button" className="text-button" onClick={() => { setForm(emptyForm); setEditingId(null) }}>ล้างข้อมูล</button><button className="primary-button" disabled={isSaving}>{isSaving ? <LoaderCircle className="spin" size={17} /> : <ArrowRight size={17} />}{editingId ? 'บันทึกการแก้ไข' : 'ลงทะเบียนชุมชน'}</button></div></div></form></section> : <section className="admin-page"><div className="admin-header"><div><div className="eyebrow"><LayoutDashboard size={15} /> แดชบอร์ดผู้ดูแล</div><h1>ภาพรวมชุมชนคุณครู</h1><p>ดูรายละเอียด สร้างภาพสรุป และส่งออกข้อมูลชุมชนทั้งหมดได้ในที่เดียว</p></div><button className="export-button" onClick={exportExcel}><FileSpreadsheet size={17} /> Export Excel</button></div><div className="admin-metrics"><Metric value={communities.length} label="ชุมชนทั้งหมด" icon={<BookOpen />} /><Metric value={communities.reduce((sum, item) => sum + item.member_count, 0)} label="สมาชิกทั้งหมด" icon={<Users />} /><Metric value="กก026" label="รหัสเริ่มต้น" icon={<KeyRound />} /></div><div className="admin-workspace"><aside className="community-list"><div className="list-head"><h3>รายชื่อชุมชน</h3><button onClick={() => beginEdit()} aria-label="เพิ่มชุมชน"><Plus size={18} /></button></div><label className="search"><Search size={17} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="ค้นหาชุมชน..." /></label><div className="list-items">{isLoading ? <div className="loading"><LoaderCircle className="spin" /> กำลังโหลด...</div> : filtered.map((community) => <button key={community.id} className={community.id === selected?.id ? 'community-item selected' : 'community-item'} onClick={() => setSelectedId(community.id)}><span className="item-avatar">{community.community_name.slice(0, 1)}</span><span><b>{community.community_name}</b><small>{community.community_code} · {community.member_count} สมาชิก</small></span></button>)}</div></aside><div className="preview-pane">{selected ? <><div className="preview-toolbar"><div><span className="preview-dot" /> ตัวอย่างภาพชุมชน</div><div><button className="icon-button" title="แก้ไข" onClick={() => beginEdit(selected)}><Edit3 size={17} /></button><button className="export-image" onClick={exportCard}><Download size={16} /> PNG</button></div></div><CommunityCard community={selected} ref={cardRef} /></> : <div className="empty">ยังไม่มีข้อมูลชุมชน</div>}</div></div></section>}
    {showLogin && <div className="modal-backdrop" onMouseDown={() => setShowLogin(false)}><section className="login-modal" onMouseDown={(event) => event.stopPropagation()}><button className="modal-close" onClick={() => setShowLogin(false)}><X size={19} /></button><div className="login-symbol"><LogIn size={25} /></div><h2>เข้าสู่ระบบ</h2><p>รับลิงก์ทางอีเมล เพื่อกลับมาแก้ไขข้อมูลชุมชนของคุณได้</p><form onSubmit={sendMagicLink}><label>อีเมล<input type="email" value={email} onChange={(event) => setEmail(event.target.value)} placeholder="you@school.ac.th" required /></label><button className="primary-button full"><LogIn size={17} /> ส่งลิงก์เข้าสู่ระบบ</button></form>{loginMessage && <div className="login-message">{loginMessage}</div>}<div className="login-divider"><span>หรือ</span></div><button className="line-button" disabled title="เชื่อมต่อ LINE Login ในขั้นตอนตั้งค่าโปรดักชัน"><span>LINE</span> เข้าสู่ระบบด้วย LINE <small>เร็ว ๆ นี้</small></button></section></div>}
  </main>
}

function Field({ label, required, className = '', children }: { label: string; required?: boolean; className?: string; children: React.ReactNode }) { return <label className={`field ${className}`}><span>{label}{required && <i>*</i>}</span>{children}</label> }
function Metric({ value, label, icon }: { value: number | string; label: string; icon: React.ReactNode }) { return <div className="metric"><span>{icon}</span><div><b>{typeof value === 'number' ? thaiNumber(value) : value}</b><small>{label}</small></div></div> }
const CommunityCard = ({ community, ref }: { community: Community; ref: React.Ref<HTMLDivElement> }) => <article className="community-card" ref={ref}><div className="card-noise" /><div className="card-left"><div className="school-badge"><span>✦</span><small>TEACHER<br />COMMUNITY</small></div><div className="card-title"><span>กิจกรรม</span><strong>ชุมนุม</strong><small>{community.school_name}</small></div><div className="photo-frame">{community.image_url ? <img src={community.image_url} alt={community.community_name} /> : <div className="photo-placeholder"><ImagePlus size={37} /><span>ภาพกิจกรรม<br />ของชุมชน</span></div>}</div></div><div className="card-right"><div className="card-row code"><span className="card-icon"><KeyRound size={16} /></span><div><label>รหัสชุมชน</label><b>{community.community_code}</b></div></div><div className="card-row"><span className="card-icon orange"><BookOpen size={16} /></span><div><label>ชื่อชุมชน</label><b>{community.community_name}</b></div></div><div className="card-row"><span className="card-icon pink"><Users size={16} /></span><div><label>ครูที่ปรึกษา</label><b>{community.advisor_name}</b></div></div><div className="card-split"><div><label><MapPin size={13} /> สถานที่</label><b>{community.location || '-'}</b></div><div><label><Users size={13} /> จำนวนที่รับ</label><b>{community.member_count || '-'}</b></div></div><div className="card-description"><label>รายละเอียด</label><p>{community.description || 'ร่วมสร้างพื้นที่เรียนรู้ที่เปี่ยมด้วยแรงบันดาลใจ'}</p></div></div></article>
export default App
