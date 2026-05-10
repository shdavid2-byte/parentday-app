import { useState, useEffect, useRef } from "react";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, set, get } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyBZR0f8bcFuiN-IGkPto0p67GXqCU70g0Q",
  authDomain: "parentday-27fc4.firebaseapp.com",
  databaseURL: "https://parentday-27fc4-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "parentday-27fc4",
  storageBucket: "parentday-27fc4.firebasestorage.app",
  messagingSenderId: "705530887879",
  appId: "1:705530887879:web:5eeccba1e045e3cda408e2"
};
const _app = initializeApp(firebaseConfig);
const _db = getDatabase(_app);

const C = {
  bg:"#F7F4F0", card:"#FFFFFF", primary:"#1E3A2F", accent:"#4E8B5F",
  warm:"#D4743A", muted:"#8A8A88", danger:"#B83232", border:"#E0D9D2",
  text:"#1A1A18", light:"#EBF4EE", purple:"#7B5EA7",
  arrived:"#E8F5E1", arrivedB:"#4E8B5F",
  cancelled:"#FDE8E8", cancelledB:"#B83232",
  noshow:"#FFF4E0", noshowB:"#D4743A",
  changed:"#EEF0FF", changedB:"#5B6FD4",
  gold:"#C9941A",
};

const pad = n => String(n).padStart(2,"0");
const toMin = t => { const [h,m]=t.split(":").map(Number); return h*60+m; };
const fromMin = m => `${pad(Math.floor(m/60))}:${pad(m%60)}`;
const uid = () => Math.random().toString(36).slice(2,9);

const genSlots = (start,end,dur) => {
  const slots=[]; let cur=toMin(start), e=toMin(end);
  while(cur+dur<=e){
    slots.push({id:uid(),start:fromMin(cur),end:fromMin(cur+dur),type:"free",booking:null,status:null,note:""});
    cur+=dur;
  }
  return slots;
};

const fmtDate = d => {
  if(!d) return "";
  const days=["ראשון","שני","שלישי","רביעי","חמישי","שישי","שבת"];
  const months=["ינואר","פברואר","מרץ","אפריל","מאי","יוני","יולי","אוגוסט","ספטמבר","אוקטובר","נובמבר","דצמבר"];
  const dt=new Date(d+"T12:00:00");
  return `יום ${days[dt.getDay()]}, ${dt.getDate()} ב${months[dt.getMonth()]} ${dt.getFullYear()}`;
};

const defaultInvite = (teacherName, subject, date, duration, link) =>
`*אסיפת הורים | מחצית ב' 📅*

שלום רב,

הגיע הזמן לשבת יחד! 🌱
אני מזמינ/ה אותך לפגישה אישית עם ילדך ואיתי — שיחה חמה ומשמעותית על ההתקדמות, החוזקות, והצעדים הבאים.

📆 *תאריך:* ${fmtDate(date)}
👤 *מורה:* ${teacherName}${subject?" ("+subject+")":""}
⏱ *משך הפגישה:* ${duration} דקות

הפגישה חשובה לקידום ילדך — ונעשה אותה יחד 💚

להרשמה לחצ/י כאן:
${link}`;

const reminderRegistered = (teacherName, subject, date, slot) =>
`*תזכורת — אסיפת הורים מחר 📅*

שלום, רצינו להזכיר לך שמחר, ${fmtDate(date)}, יש לך שיחה עם ${teacherName}${subject?" ("+subject+")":""}.
🕐 *שעה:* ${slot.start}–${slot.end}
נשמח לראותך! 😊`;

const reminderUnregistered = (teacherName, subject, date, link, studentName) =>
`*תזכורת הרשמה — אסיפת הורים 📅*

שלום, שמנו לב שטרם נרשמת לשיחת הורים עם ${teacherName}${subject?" ("+subject+")":""} בנוגע ל${studentName}.
📆 *תאריך:* ${fmtDate(date)}

עדיין יש שעות פנויות — להרשמה:
${link}

נשמח לראותך! 💚`;

const store = {
  async get(k){
    try{
      const safeKey = k.replace(/[.#$\[\]]/g,'_');
      const snapshot = await get(ref(_db, safeKey));
      return snapshot.exists() ? snapshot.val() : null;
    }catch(e){ console.error('Firebase get:',e); return null; }
  },
  async set(k,v){
    try{
      const safeKey = k.replace(/[.#$\[\]]/g,'_');
      await set(ref(_db, safeKey), v);
      return true;
    }catch(e){ console.error('Firebase set:',e); return false; }
  },
};

const parseStudents = (text) =>
  text.split("\n").map(l=>l.trim()).filter(Boolean).map(line=>{
    const parts=line.split(/\t|,|  +/).map(p=>p.trim()).filter(Boolean);
    const phoneIdx=parts.findIndex(p=>/^[\d\-\+\s]{7,}$/.test(p));
    if(parts.length>=2&&phoneIdx>=0){
      const name=parts.filter((_,i)=>i!==phoneIdx).join(" ");
      return {id:uid(),name,phone:parts[phoneIdx]};
    }
    return {id:uid(),name:parts.join(" "),phone:""};
  });

const exportXLSX = async (data) => {
  const XLSX=await import("https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs");
  const wb=XLSX.utils.book_new();
  const statusLabel={arrived:"הגיע/ה",cancelled:"ביטל/ה",noshow:"לא הגיע/ה",changed:"שונה שעה"};
  const rows1=[
    [`סיכום יום הורים — ${data.teacherName}${data.subject?" ("+data.subject+")":""}`],
    [`תאריך: ${fmtDate(data.date)}`],[],
    ["שעה","עד","שם הורה","שם ילד/ה","טלפון","סטטוס","הערה"],
  ];
  data.slots.forEach(s=>{
    if(s.type==="break"){rows1.push([s.start,s.end,"—הפסקה—","","","",""]);return;}
    if(!s.booking){rows1.push([s.start,s.end,"פנוי","","","",""]);return;}
    rows1.push([s.start,s.end,s.booking.parentName,s.booking.childName,s.booking.phone||"",statusLabel[s.status]||"נרשם/ה",s.note||""]);
  });
  const booked=data.slots.filter(s=>s.booking).length;
  const arrived=data.slots.filter(s=>s.status==="arrived").length;
  const cancelled=data.slots.filter(s=>s.status==="cancelled").length;
  const noshow=data.slots.filter(s=>s.status==="noshow").length;
  rows1.push([],[],["סיכום:","נרשמו:",booked,"הגיעו:",arrived,"ביטלו:",cancelled,"לא הגיעו:",noshow]);
  const ws1=XLSX.utils.aoa_to_sheet(rows1);
  ws1["!cols"]=[{wch:10},{wch:10},{wch:20},{wch:18},{wch:14},{wch:12},{wch:24}];
  XLSX.utils.book_append_sheet(wb,ws1,"לוח שעות");
  if(data.students?.length){
    const regNames=data.slots.filter(s=>s.booking).map(s=>s.booking.childName.trim().toLowerCase());
    const unreg=data.students.filter(s=>!regNames.includes(s.name.trim().toLowerCase()));
    const rows2=[["תלמידים שלא נרשמו"],[],["שם תלמיד","טלפון הורה"],...unreg.map(s=>[s.name,s.phone||""])];
    const ws2=XLSX.utils.aoa_to_sheet(rows2);
    ws2["!cols"]=[{wch:22},{wch:16}];
    XLSX.utils.book_append_sheet(wb,ws2,"לא נרשמו");
  }
  XLSX.writeFile(wb,`יום_הורים_${data.teacherName}_${data.date||"לוח"}.xlsx`);
};

// ── UI ────────────────────────────────────────────────────────────────────────
const iStyle={width:"100%",padding:"9px 12px",border:`1.5px solid ${C.border}`,borderRadius:8,fontSize:14,background:C.card,color:C.text,boxSizing:"border-box",fontFamily:"inherit",outline:"none"};
function Lbl({children}){ return <label style={{display:"block",fontSize:11,fontWeight:700,color:C.primary,marginBottom:5,textTransform:"uppercase",letterSpacing:.6}}>{children}</label>; }
function Btn({onClick,disabled,color=C.primary,ghost,full,sm,danger,children,style:sx={}}){
  return <button onClick={onClick} disabled={disabled} style={{padding:sm?"6px 13px":"10px 20px",width:full?"100%":undefined,background:ghost?"transparent":disabled?C.border:danger?C.danger:color,color:ghost?(danger?C.danger:color):"white",border:ghost?`1.5px solid ${danger?C.danger:C.border}`:"none",borderRadius:9,fontWeight:700,fontSize:sm?13:14,cursor:disabled?"not-allowed":"pointer",fontFamily:"inherit",transition:"all .15s",opacity:disabled?.6:1,...sx}}>{children}</button>;
}
function Tag({color,children}){ return <span style={{padding:"3px 10px",borderRadius:20,background:color+"22",color,fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>{children}</span>; }
function Modal({onClose,children,wide}){
  return(
    <div style={{position:"fixed",inset:0,background:"#0009",display:"flex",alignItems:"center",justifyContent:"center",zIndex:400,padding:16}} onClick={onClose}>
      <div style={{background:C.card,borderRadius:18,padding:26,width:"100%",maxWidth:wide?520:440,boxShadow:"0 28px 70px #0005",maxHeight:"92vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>{children}</div>
    </div>
  );
}

function Footer(){
  return(
    <div style={{textAlign:"center",padding:"20px 16px 32px",borderTop:`1px solid ${C.border}`,marginTop:32}}>
      <div style={{fontFamily:"Georgia, serif",fontSize:13,color:C.muted,letterSpacing:"0.08em"}}>
        <span style={{fontWeight:700,color:C.primary,fontSize:14}}>© SD</span>
        <span style={{margin:"0 6px",color:C.border}}>|</span>
        <span style={{fontSize:12}}>שירן דוד · 2026</span>
        <span style={{margin:"0 6px",color:C.border}}>|</span>
        <span style={{fontSize:11,fontStyle:"italic"}}>כל הזכויות שמורות</span>
      </div>
    </div>
  );
}

// ── WELCOME / INSTRUCTIONS ────────────────────────────────────────────────────
function Welcome({onStart}){
  const steps=[
    {icon:"⚙️", title:"הגדר את היום", desc:"הכנס שם, תאריך, שעות ומשך כל שיחה"},
    {icon:"📋", title:"הוסף רשימת תלמידים", desc:"הדבק שמות וטלפונים כדי לעקוב מי לא נרשם"},
    {icon:"🔐", title:"שמור את קישור הניהול", desc:"קישור קבוע שרק אתה/את שומרת — לניהול ומעקב"},
    {icon:"📤", title:"שלח להורים", desc:"קישור הרשמה — הורים בוחרים שעה בעצמם"},
    {icon:"📊", title:"ביום עצמו", desc:"סמן הגיע / ביטל / לא הגיע, ובסוף ייצא לאקסל"},
  ];

  return(
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Heebo','Arial Hebrew',Arial,sans-serif",direction:"rtl",color:C.text}}>
      <div style={{background:C.primary,padding:"28px 20px",textAlign:"center"}}>
        <div style={{width:72,height:72,background:"rgba(255,255,255,0.15)",borderRadius:20,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:36,marginBottom:14}}>📅</div>
        <h1 style={{color:"white",margin:"0 0 6px",fontSize:26,fontWeight:900}}>יום הורים</h1>
        <p style={{color:"#A8D5B5",margin:0,fontSize:15}}>מערכת הרשמה חכמה לשיחות הורים</p>
      </div>

      <div style={{maxWidth:480,margin:"0 auto",padding:"28px 16px"}}>

        <div style={{background:C.card,borderRadius:16,border:`1.5px solid ${C.border}`,padding:22,marginBottom:20}}>
          <h2 style={{margin:"0 0 16px",fontSize:17,color:C.primary,fontWeight:800}}>איך זה עובד? 👇</h2>
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            {steps.map((s,i)=>(
              <div key={i} style={{display:"flex",gap:14,alignItems:"flex-start"}}>
                <div style={{width:38,height:38,borderRadius:10,background:C.light,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{s.icon}</div>
                <div>
                  <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:2}}>{s.title}</div>
                  <div style={{fontSize:13,color:C.muted,lineHeight:1.5}}>{s.desc}</div>
                </div>
              </div>
            ))}
          </div>
        </div>

        <div style={{background:"#FFF8E7",border:`1.5px solid ${C.gold}44`,borderRadius:12,padding:"14px 16px",marginBottom:24}}>
          <div style={{fontSize:13,color:C.gold,fontWeight:700,marginBottom:4}}>💡 חשוב לדעת</div>
          <div style={{fontSize:13,color:C.text,lineHeight:1.7}}>
            בסוף ההגדרה תקבל/י <strong>קישור ניהול קבוע</strong> — שמור/י אותו! זה הקישור שלך לכל הנתונים.<br/>
            ההורים מקבלים קישור נפרד <strong>להרשמה בלבד</strong>.
          </div>
        </div>

        <Btn full onClick={onStart} style={{fontSize:17,padding:"15px",fontWeight:900}}>
          🚀 התחל/י עכשיו
        </Btn>
      </div>
      <Footer/>
    </div>
  );
}

// ── SAVE LINKS SCREEN ─────────────────────────────────────────────────────────
function SaveLinks({data,sessionId,onDone}){
  const baseUrl=window.location.href.split("?")[0];
  const adminLink=`${baseUrl}?admin=${sessionId}`;
  const parentLink=`${baseUrl}?s=${sessionId}`;
  const [copiedA,setCopiedA]=useState(false);
  const [copiedP,setCopiedP]=useState(false);

  const copy=(text,setCopied)=>navigator.clipboard.writeText(text).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);});

  const selfMsg=`*קישור ניהול — יום הורים שלי 🔐*\n\n${data.teacherName}${data.subject?" ("+data.subject+")":""}\n${fmtDate(data.date)}\n\nקישור הניהול שלי (לשמור!):\n${adminLink}\n\nקישור להורים:\n${parentLink}`;

  return(
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Heebo','Arial Hebrew',Arial,sans-serif",direction:"rtl",color:C.text,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{width:"100%",maxWidth:460}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{fontSize:56,marginBottom:8}}>🎉</div>
          <h1 style={{fontSize:22,fontWeight:900,color:C.primary,margin:"0 0 6px"}}>יום ההורים שלך מוכן!</h1>
          <p style={{color:C.muted,fontSize:14,margin:0}}>שמור/י את הקישורים הבאים לפני שממשיכים</p>
        </div>

        {/* Admin link */}
        <div style={{background:C.card,borderRadius:14,border:`2px solid ${C.primary}`,padding:20,marginBottom:14}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <span style={{fontSize:20}}>🔐</span>
            <div>
              <div style={{fontWeight:800,fontSize:15,color:C.primary}}>קישור ניהול — שלך בלבד!</div>
              <div style={{fontSize:12,color:C.muted}}>פתח/י אותו בכל זמן כדי לראות הרשמות ולנהל</div>
            </div>
          </div>
          <div style={{background:C.light,borderRadius:8,padding:"8px 12px",fontSize:11,color:C.muted,wordBreak:"break-all",direction:"ltr",marginBottom:10}}>{adminLink}</div>
          <div style={{display:"flex",gap:8}}>
            <Btn sm full color={copiedA?C.accent:C.primary} onClick={()=>copy(adminLink,setCopiedA)}>{copiedA?"✓ הועתק":"📋 העתק"}</Btn>
            <a href={`https://wa.me/?text=${encodeURIComponent(selfMsg)}`} target="_blank" rel="noopener noreferrer"
              style={{flex:1,display:"block",padding:"6px 13px",background:"#25D366",color:"white",textDecoration:"none",borderRadius:9,fontWeight:700,fontSize:13,textAlign:"center"}}>
              📱 שלח לעצמי
            </a>
          </div>
        </div>

        {/* Parent link */}
        <div style={{background:C.card,borderRadius:14,border:`1.5px solid ${C.border}`,padding:20,marginBottom:20}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <span style={{fontSize:20}}>👪</span>
            <div>
              <div style={{fontWeight:800,fontSize:15,color:C.text}}>קישור להורים</div>
              <div style={{fontSize:12,color:C.muted}}>שלח/י לאחר מכן דרך לוח הניהול</div>
            </div>
          </div>
          <div style={{background:C.bg,borderRadius:8,padding:"8px 12px",fontSize:11,color:C.muted,wordBreak:"break-all",direction:"ltr",marginBottom:10}}>{parentLink}</div>
          <Btn sm full ghost onClick={()=>copy(parentLink,setCopiedP)}>{copiedP?"✓ הועתק":"📋 העתק"}</Btn>
        </div>

        <div style={{background:"#FFF8E7",border:`1px solid ${C.gold}44`,borderRadius:10,padding:"12px 16px",marginBottom:20,fontSize:13,color:C.text,lineHeight:1.7}}>
          ⚠️ <strong>שמור/י את קישור הניהול!</strong> בלעדיו לא תוכל/י לחזור לנתונים. שלח/י אותו לעצמך בוואטסאפ עכשיו.
        </div>

        <Btn full onClick={onDone} style={{fontSize:16,padding:"14px"}}>כניסה לניהול יום ההורים ←</Btn>
        <Footer/>
      </div>
    </div>
  );
}

// ── STUDENT ROSTER ────────────────────────────────────────────────────────────
function StudentRoster({onDone}){
  const [mode,setMode]=useState("paste");
  const [text,setText]=useState("");
  const [preview,setPreview]=useState([]);
  const [error,setError]=useState("");
  const fileRef=useRef();

  const parseAndPreview=(t)=>{
    const parsed=parseStudents(t);
    setPreview(parsed);
    setError(parsed.length===0?"לא זוהו תלמידים — בדוק/י את הפורמט":"");
  };

  const handleFile=async(e)=>{
    const file=e.target.files[0]; if(!file) return;
    const ext=file.name.split(".").pop().toLowerCase();
    if(ext==="csv"||ext==="txt"){ const t=await file.text(); setText(t); parseAndPreview(t); }
    else if(ext==="xlsx"||ext==="xls"){
      try{
        const XLSX=await import("https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs");
        const buf=await file.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1});
        const t=rows.filter(r=>r.length>=1).map(r=>r.join("\t")).join("\n");
        setText(t); parseAndPreview(t);
      }catch(err){setError("שגיאה: "+err.message);}
    } else { setError("פורמט לא נתמך — CSV, XLSX, או TXT"); }
  };

  return(
    <div style={{background:C.card,borderRadius:16,border:`1.5px solid ${C.border}`,padding:22}}>
      <h3 style={{margin:"0 0 4px",color:C.primary,fontSize:16}}>📋 רשימת תלמידים</h3>
      <p style={{margin:"0 0 16px",color:C.muted,fontSize:13}}>אופציונלי — לשליחת תזכורות למי שלא נרשם</p>
      <div style={{display:"flex",gap:0,marginBottom:16,background:C.border+"55",borderRadius:9,padding:3}}>
        {[["paste","📝 הדבקה"],["upload","📂 קובץ"]].map(([m,l])=>(
          <button key={m} onClick={()=>setMode(m)} style={{flex:1,padding:"7px",borderRadius:7,border:"none",background:mode===m?C.card:"transparent",color:mode===m?C.primary:C.muted,fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>{l}</button>
        ))}
      </div>
      {mode==="paste"&&(
        <div>
          <Lbl>שם + טלפון הורה (שורה לכל תלמיד)</Lbl>
          <div style={{fontSize:11,color:C.muted,marginBottom:6}}>דוגמה: דניאל כהן, 050-1234567</div>
          <textarea style={{...iStyle,height:150,resize:"vertical",fontSize:13,lineHeight:1.8}} placeholder={"דניאל כהן, 050-1234567\nנועה לוי\t052-9876543"} value={text} onChange={e=>{setText(e.target.value);parseAndPreview(e.target.value);}}/>
        </div>
      )}
      {mode==="upload"&&(
        <div onClick={()=>fileRef.current.click()} style={{border:`2px dashed ${C.accent}`,borderRadius:12,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:C.light}}>
          <div style={{fontSize:32,marginBottom:8}}>📂</div>
          <div style={{fontWeight:700,color:C.primary,marginBottom:4}}>לחץ/י להעלאת קובץ</div>
          <div style={{fontSize:12,color:C.muted}}>CSV, XLSX, TXT · עמודות: שם, טלפון</div>
          <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,.txt" style={{display:"none"}} onChange={handleFile}/>
        </div>
      )}
      {error&&<div style={{color:C.danger,fontSize:13,marginTop:8}}>⚠️ {error}</div>}
      {preview.length>0&&(
        <div style={{marginTop:14}}>
          <div style={{fontSize:12,fontWeight:700,color:C.accent,marginBottom:8}}>✓ זוהו {preview.length} תלמידים</div>
          <div style={{maxHeight:150,overflowY:"auto",display:"flex",flexDirection:"column",gap:5}}>
            {preview.slice(0,6).map((s,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 10px",borderRadius:7,background:C.light,fontSize:13}}>
                <span style={{fontWeight:600}}>{s.name}</span>
                <span style={{color:C.muted}}>{s.phone||"ללא טלפון"}</span>
              </div>
            ))}
            {preview.length>6&&<div style={{fontSize:12,color:C.muted,textAlign:"center",padding:"4px 0"}}>ועוד {preview.length-6}...</div>}
          </div>
          <Btn full style={{marginTop:12}} onClick={()=>onDone(preview)}>✓ אשר רשימה ({preview.length})</Btn>
        </div>
      )}
      <button onClick={()=>onDone([])} style={{background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:13,fontFamily:"inherit",marginTop:12,width:"100%",textAlign:"center"}}>דלג/י — ללא רשימת תלמידים</button>
    </div>
  );
}

// ── SETUP ─────────────────────────────────────────────────────────────────────
function Setup({onComplete}){
  const [name,setName]=useState(""); const [subject,setSubject]=useState("");
  const [date,setDate]=useState(""); const [start,setStart]=useState("08:00");
  const [end,setEnd]=useState("14:00"); const [dur,setDur]=useState(10);
  const [step,setStep]=useState("info");
  const [saving,setSaving]=useState(false);
  const canNext=name.trim()&&date&&toMin(end)>toMin(start)+dur;

  const go=async(students)=>{
    setSaving(true);
    const sid=uid().toUpperCase();
    const slots=genSlots(start,end,dur);
    const d={teacherName:name.trim(),subject,date,startTime:start,endTime:end,duration:dur,slots,students,sid};
    await store.set(`session:${sid}`,d);
    await store.set(`admin:${sid}`,{sid,created:new Date().toISOString()});
    onComplete(d);
    setSaving(false);
  };

  return(
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",padding:20,fontFamily:"'Heebo','Arial Hebrew',Arial,sans-serif",direction:"rtl"}}>
      <div style={{width:"100%",maxWidth:480}}>
        <div style={{textAlign:"center",marginBottom:20}}>
          <h1 style={{fontSize:22,fontWeight:900,color:C.primary,margin:"0 0 4px"}}>הגדרת יום הורים</h1>
          <div style={{display:"flex",alignItems:"center",gap:8,justifyContent:"center",marginTop:12}}>
            {[["info","1","פרטים"],["students","2","תלמידים"]].map(([s,n,l],i)=>(
              <div key={s} style={{display:"flex",alignItems:"center",gap:8}}>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <div style={{width:26,height:26,borderRadius:"50%",background:step===s?C.primary:step==="students"&&s==="info"?C.accent:C.border+"99",color:step===s||(step==="students"&&s==="info")?"white":C.muted,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:12,transition:"all .2s"}}>{step==="students"&&s==="info"?"✓":n}</div>
                  <span style={{fontSize:13,fontWeight:600,color:step===s?C.primary:C.muted}}>{l}</span>
                </div>
                {i===0&&<div style={{width:20,height:2,background:C.border}}/>}
              </div>
            ))}
          </div>
        </div>

        {step==="info"&&(
          <div style={{background:C.card,borderRadius:16,border:`1.5px solid ${C.border}`,padding:22,display:"flex",flexDirection:"column",gap:16}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
              <div><Lbl>שם המורה *</Lbl><input style={iStyle} placeholder="שם מלא..." value={name} onChange={e=>setName(e.target.value)}/></div>
              <div><Lbl>מקצוע</Lbl><input style={iStyle} placeholder="מתמטיקה..." value={subject} onChange={e=>setSubject(e.target.value)}/></div>
            </div>
            <div><Lbl>תאריך *</Lbl><input style={iStyle} type="date" value={date} onChange={e=>setDate(e.target.value)}/></div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
              <div><Lbl>שעת התחלה</Lbl><input style={iStyle} type="time" value={start} onChange={e=>setStart(e.target.value)}/></div>
              <div><Lbl>שעת סיום</Lbl><input style={iStyle} type="time" value={end} onChange={e=>setEnd(e.target.value)}/></div>
            </div>
            <div>
              <Lbl>משך כל שיחה (דקות)</Lbl>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {[5,8,10,12,15,20].map(d=>(
                  <button key={d} onClick={()=>setDur(d)} style={{padding:"7px 16px",borderRadius:8,border:`2px solid ${dur===d?C.primary:C.border}`,background:dur===d?C.primary:"white",color:dur===d?"white":C.text,fontWeight:700,cursor:"pointer",fontSize:14,fontFamily:"inherit"}}>{d}′</button>
                ))}
              </div>
            </div>
            <Btn disabled={!canNext} onClick={()=>setStep("students")} full style={{fontSize:15,padding:"12px",marginTop:4}}>המשך לרשימת תלמידים →</Btn>
          </div>
        )}

        {step==="students"&&(
          <>
            <StudentRoster onDone={go}/>
            {saving&&<div style={{textAlign:"center",marginTop:12,color:C.muted,fontSize:14}}>⏳ יוצר יום הורים...</div>}
          </>
        )}
        <Footer/>
      </div>
    </div>
  );
}

// ── PARENT VIEW ───────────────────────────────────────────────────────────────
function ParentView({data,onBook}){
  const [parentName,setParentName]=useState("");
  const [phone,setPhone]=useState("");
  const [childName,setChildName]=useState("");
  const [chosenSlot,setChosenSlot]=useState(null);
  const [step,setStep]=useState("form");
  const [saving,setSaving]=useState(false);
  const freeSlots=data.slots.filter(s=>s.type==="free"&&!s.booking);
  const canNext=parentName.trim()&&phone.trim()&&childName.trim();

  const book=async()=>{ setSaving(true); await onBook(chosenSlot.id,{parentName:parentName.trim(),phone:phone.trim(),childName:childName.trim()}); setSaving(false); setStep("done"); };

  return(
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Heebo','Arial Hebrew',Arial,sans-serif",direction:"rtl",color:C.text}}>
      <div style={{background:C.primary,padding:"14px 20px"}}>
        <div style={{maxWidth:480,margin:"0 auto"}}>
          <h1 style={{color:"white",margin:0,fontSize:17,fontWeight:800}}>📅 {data.teacherName}{data.subject?` · ${data.subject}`:""}</h1>
          <div style={{color:"#A8D5B5",fontSize:13,marginTop:2}}>{fmtDate(data.date)} · {data.duration}′ לשיחה</div>
        </div>
      </div>
      <div style={{maxWidth:480,margin:"0 auto",padding:"24px 16px"}}>
        {step==="done"&&(
          <div style={{background:C.card,borderRadius:16,border:`1.5px solid ${C.border}`,padding:"40px 24px",textAlign:"center"}}>
            <div style={{fontSize:56,marginBottom:12}}>✅</div>
            <h2 style={{color:C.primary,margin:"0 0 8px"}}>נרשמת בהצלחה!</h2>
            <p style={{color:C.muted,fontSize:15,margin:"0 0 20px"}}>{childName} · {chosenSlot.start}–{chosenSlot.end}</p>
            <div style={{background:C.light,borderRadius:10,padding:"14px 16px",fontSize:14,color:C.primary,lineHeight:1.9}}>
              📍 {fmtDate(data.date)}<br/>🕐 {chosenSlot.start}–{chosenSlot.end}<br/>👤 שיחה עם {data.teacherName}
            </div>
          </div>
        )}
        {step==="form"&&(
          <>
            <div style={{marginBottom:20}}>
              <h2 style={{fontSize:18,fontWeight:800,color:C.primary,margin:"0 0 4px"}}>הרשמה לשיחת הורים</h2>
              <p style={{color:C.muted,fontSize:14,margin:0}}>מלא/י את הפרטים ובחר/י שעה</p>
            </div>
            <div style={{background:C.card,borderRadius:14,border:`1.5px solid ${C.border}`,padding:20,display:"flex",flexDirection:"column",gap:14,marginBottom:16}}>
              <div><Lbl>שם ההורה *</Lbl><input style={iStyle} placeholder="שם מלא..." value={parentName} onChange={e=>setParentName(e.target.value)}/></div>
              <div><Lbl>מספר פלאפון *</Lbl><input style={iStyle} type="tel" placeholder="050-0000000" value={phone} onChange={e=>setPhone(e.target.value)}/></div>
              <div><Lbl>שם הילד/ה *</Lbl><input style={iStyle} placeholder="שם הילד/ה..." value={childName} onChange={e=>setChildName(e.target.value)}/></div>
            </div>
            <Btn full disabled={!canNext} onClick={()=>setStep("slot")}>המשך לבחירת שעה →</Btn>
          </>
        )}
        {step==="slot"&&(
          <>
            <button onClick={()=>setStep("form")} style={{background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:14,marginBottom:16,fontFamily:"inherit",padding:0}}>← חזרה לפרטים</button>
            <div style={{marginBottom:16}}>
              <h2 style={{fontSize:18,fontWeight:800,color:C.primary,margin:"0 0 2px"}}>בחר/י שעה</h2>
              <div style={{color:C.muted,fontSize:14}}>{parentName} · {childName}</div>
            </div>
            {freeSlots.length===0
              ?<div style={{background:C.card,borderRadius:14,border:`1.5px solid ${C.border}`,padding:32,textAlign:"center",color:C.muted}}><div style={{fontSize:36,marginBottom:10}}>😕</div>אין שעות פנויות</div>
              :<>
                <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:20}}>
                  {freeSlots.map(sl=>(
                    <button key={sl.id} onClick={()=>setChosenSlot(sl)} style={{padding:"13px 18px",borderRadius:12,border:`2px solid ${chosenSlot?.id===sl.id?C.primary:C.border}`,background:chosenSlot?.id===sl.id?C.light:"white",cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"space-between",transition:"all .15s"}}>
                      <span style={{fontSize:16,fontWeight:700,direction:"ltr"}}>{sl.start} – {sl.end}</span>
                      {chosenSlot?.id===sl.id&&<span style={{color:C.primary,fontSize:18}}>✓</span>}
                    </button>
                  ))}
                </div>
                <Btn full disabled={!chosenSlot||saving} onClick={book}>{saving?"⏳ שומר...":"✓ אשר הרשמה — "+(chosenSlot?.start||"")}</Btn>
              </>
            }
          </>
        )}
      </div>
      <Footer/>
    </div>
  );
}

// ── SLOT MODAL ────────────────────────────────────────────────────────────────
function SlotModal({slot,allSlots,onClose,onSave}){
  const [status,setStatus]=useState(slot.status||"");
  const [note,setNote]=useState(slot.note||"");
  const [moveTo,setMoveTo]=useState("");
  const [tab,setTab]=useState("status");
  const freeSlots=allSlots.filter(s=>s.id!==slot.id&&s.type==="free"&&!s.booking);
  const opts=[{v:"arrived",label:"✅ הגיע/ה",color:C.accent},{v:"cancelled",label:"❌ ביטל/ה",color:C.danger},{v:"noshow",label:"⏳ לא הגיע/ה",color:C.warm}];
  return(
    <Modal onClose={onClose}>
      <div style={{marginBottom:16}}>
        <div style={{fontWeight:800,fontSize:17,color:C.primary}}>{slot.start} – {slot.end}</div>
        {slot.booking&&<div style={{color:C.muted,fontSize:14,marginTop:2}}>{slot.booking.parentName} · {slot.booking.childName}{slot.booking.phone&&` · ${slot.booking.phone}`}</div>}
      </div>
      <div style={{display:"flex",gap:0,marginBottom:16,background:C.border+"55",borderRadius:9,padding:3}}>
        {[["status","סטטוס"],["move","שינוי שעה"],["note","הערה"]].map(([t,l])=>(
          <button key={t} onClick={()=>setTab(t)} style={{flex:1,padding:"7px",borderRadius:7,border:"none",background:tab===t?C.card:"transparent",color:tab===t?C.primary:C.muted,fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>{l}</button>
        ))}
      </div>
      {tab==="status"&&<div style={{display:"flex",flexDirection:"column",gap:8}}>{opts.map(o=><button key={o.v} onClick={()=>setStatus(status===o.v?"":o.v)} style={{padding:"12px 16px",borderRadius:10,border:`2px solid ${status===o.v?o.color:C.border}`,background:status===o.v?o.color+"18":"white",cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:15,color:status===o.v?o.color:C.text,textAlign:"right",transition:"all .15s"}}>{o.label}</button>)}</div>}
      {tab==="move"&&(freeSlots.length===0?<div style={{color:C.muted,textAlign:"center",padding:"20px 0",fontSize:14}}>אין שעות פנויות</div>:<div style={{display:"flex",flexDirection:"column",gap:8}}>{freeSlots.map(sl=><button key={sl.id} onClick={()=>setMoveTo(moveTo===sl.id?"":sl.id)} style={{padding:"11px 16px",borderRadius:10,border:`2px solid ${moveTo===sl.id?C.primary:C.border}`,background:moveTo===sl.id?C.light:"white",cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:15,direction:"ltr",textAlign:"center",transition:"all .15s"}}>{sl.start} – {sl.end}</button>)}</div>)}
      {tab==="note"&&<textarea style={{...iStyle,height:100,resize:"vertical"}} placeholder="הוסף הערה..." value={note} onChange={e=>setNote(e.target.value)}/>}
      <div style={{display:"flex",gap:10,marginTop:18}}>
        <Btn ghost full onClick={onClose}>ביטול</Btn>
        <Btn full onClick={()=>onSave({status:status||null,note,moveTo:moveTo||null})}>💾 שמור</Btn>
      </div>
    </Modal>
  );
}

// ── SHARE MODAL ───────────────────────────────────────────────────────────────
function ShareModal({data,sessionId,onClose}){
  const baseUrl=window.location.href.split("?")[0];
  const link=`${baseUrl}?s=${sessionId}`;
  const [msg,setMsg]=useState(defaultInvite(data.teacherName,data.subject,data.date,data.duration,link));
  const [editing,setEditing]=useState(false);
  const [copied,setCopied]=useState(false);
  const copy=()=>navigator.clipboard.writeText(link).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);});
  return(
    <Modal onClose={onClose}>
      <h3 style={{margin:"0 0 4px",color:C.primary}}>📤 שלח הזמנה להורים</h3>
      <p style={{margin:"0 0 14px",color:C.muted,fontSize:13}}>הודעה מנוסחת מראש — ניתן לערוך</p>
      <div style={{position:"relative",marginBottom:14}}>
        {editing?<textarea value={msg} onChange={e=>setMsg(e.target.value)} style={{...iStyle,height:220,resize:"vertical",fontSize:13,lineHeight:1.7}}/>
          :<div style={{background:"#F0F7F0",border:`1.5px solid ${C.accent}44`,borderRadius:10,padding:"12px 14px",fontSize:13,color:C.text,lineHeight:1.8,whiteSpace:"pre-wrap"}}>{msg}</div>}
        <button onClick={()=>setEditing(e=>!e)} style={{position:"absolute",top:8,left:8,background:C.card,border:`1px solid ${C.border}`,borderRadius:7,padding:"3px 10px",fontSize:12,cursor:"pointer",color:C.muted,fontFamily:"inherit",fontWeight:600}}>{editing?"✓ סיים":"✏️ ערוך"}</button>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        <a href={`https://wa.me/?text=${encodeURIComponent(msg)}`} target="_blank" rel="noopener noreferrer" style={{display:"block",padding:"12px",background:"#25D366",color:"white",textDecoration:"none",borderRadius:9,fontWeight:700,fontSize:15,textAlign:"center"}}>📱 שלח בוואטסאפ</a>
        <Btn full color={copied?C.accent:C.primary} onClick={copy}>{copied?"✓ הועתק!":"📋 העתק קישור"}</Btn>
        <button onClick={()=>setMsg(defaultInvite(data.teacherName,data.subject,data.date,data.duration,link))} style={{background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>↺ שחזר הודעה מקורית</button>
        <Btn ghost full onClick={onClose}>סגור</Btn>
      </div>
    </Modal>
  );
}

// ── REMINDERS MODAL ───────────────────────────────────────────────────────────
function RemindersModal({data,sessionId,onClose}){
  const baseUrl=window.location.href.split("?")[0];
  const link=`${baseUrl}?s=${sessionId}`;
  const regWithPhone=data.slots.filter(s=>s.booking?.phone&&s.type!=="break");
  const regNames=data.slots.filter(s=>s.booking).map(s=>s.booking.childName.trim().toLowerCase());
  const unreg=(data.students||[]).filter(s=>!regNames.includes(s.name.trim().toLowerCase()));
  const unregWithPhone=unreg.filter(s=>s.phone);
  const unregNoPhone=unreg.filter(s=>!s.phone);
  const toIntl=p=>{ const n=p.replace(/[-\s]/g,""); return n.startsWith("0")?"972"+n.slice(1):n; };
  const [tab,setTab]=useState("registered");
  return(
    <Modal onClose={onClose} wide>
      <h3 style={{margin:"0 0 4px",color:C.primary}}>🔔 שליחת תזכורות</h3>
      <p style={{margin:"0 0 14px",color:C.muted,fontSize:13}}>לחץ/י על הורה לשליחת תזכורת אישית בוואטסאפ</p>
      <div style={{display:"flex",gap:0,marginBottom:16,background:C.border+"55",borderRadius:9,padding:3}}>
        {[["registered",`נרשמו (${regWithPhone.length})`],["unreg",`לא נרשמו (${unreg.length})`]].map(([t,l])=>(
          <button key={t} onClick={()=>setTab(t)} style={{flex:1,padding:"7px",borderRadius:7,border:"none",background:tab===t?C.card:"transparent",color:tab===t?C.primary:C.muted,fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>{l}</button>
        ))}
      </div>
      {tab==="registered"&&(regWithPhone.length===0?<div style={{textAlign:"center",padding:"28px 0",color:C.muted}}>אין הרשמות עם טלפון</div>:
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {regWithPhone.map(sl=>(
            <a key={sl.id} href={`https://wa.me/${toIntl(sl.booking.phone)}?text=${encodeURIComponent(reminderRegistered(data.teacherName,data.subject,data.date,sl))}`} target="_blank" rel="noopener noreferrer"
              style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 14px",borderRadius:10,border:`1.5px solid ${C.border}`,background:C.card,textDecoration:"none",color:C.text}}>
              <div>
                <div style={{fontWeight:700,fontSize:14}}>{sl.booking.parentName} · <span style={{color:C.muted,fontWeight:400}}>{sl.booking.childName}</span></div>
                <div style={{fontSize:12,color:C.muted,marginTop:2,direction:"ltr"}}>{sl.start} · {sl.booking.phone}</div>
              </div>
              <span style={{background:"#25D366",color:"white",borderRadius:8,padding:"5px 12px",fontSize:13,fontWeight:700,whiteSpace:"nowrap"}}>📱 שלח</span>
            </a>
          ))}
        </div>
      )}
      {tab==="unreg"&&(unreg.length===0?<div style={{textAlign:"center",padding:"28px 0",color:C.muted}}><div style={{fontSize:36,marginBottom:8}}>🎉</div>כל התלמידים נרשמו!</div>:
        <>
          {unregWithPhone.map(s=>(
            <a key={s.id} href={`https://wa.me/${toIntl(s.phone)}?text=${encodeURIComponent(reminderUnregistered(data.teacherName,data.subject,data.date,link,s.name))}`} target="_blank" rel="noopener noreferrer"
              style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 14px",borderRadius:10,border:`1.5px solid ${C.noshowB}44`,background:C.noshow,textDecoration:"none",color:C.text,marginBottom:8}}>
              <div>
                <div style={{fontWeight:700,fontSize:14}}>{s.name}</div>
                <div style={{fontSize:12,color:C.muted,marginTop:2}}>{s.phone}</div>
              </div>
              <span style={{background:"#25D366",color:"white",borderRadius:8,padding:"5px 12px",fontSize:13,fontWeight:700,whiteSpace:"nowrap"}}>📱 שלח</span>
            </a>
          ))}
          {unregNoPhone.length>0&&<div style={{background:C.border+"44",borderRadius:10,padding:"10px 14px",marginTop:8}}><div style={{fontSize:12,fontWeight:700,color:C.muted,marginBottom:6}}>ללא טלפון:</div>{unregNoPhone.map(s=><div key={s.id} style={{fontSize:13,color:C.muted,marginBottom:2}}>• {s.name}</div>)}</div>}
        </>
      )}
      <Btn ghost full onClick={onClose} style={{marginTop:16}}>סגור</Btn>
    </Modal>
  );
}

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
function Dashboard({data,onUpdate,sessionId}){
  const [tab,setTab]=useState("schedule");
  const [activeSlot,setActiveSlot]=useState(null);
  const [shareOpen,setShareOpen]=useState(false);
  const [remindOpen,setRemindOpen]=useState(false);
  const [exporting,setExporting]=useState(false);

  const save=async(nd)=>{ await store.set(`session:${sessionId}`,nd); onUpdate(nd); };
  const toggleBreak=async(id)=>{ const slots=data.slots.map(s=>s.id!==id?s:{...s,type:s.type==="break"?"free":"break",booking:null,status:null,note:""}); await save({...data,slots}); };
  const cancelSlot=async(id)=>{ const slots=data.slots.map(s=>s.id!==id?s:{...s,booking:null,type:"free",status:null,note:""}); await save({...data,slots}); };
  const handleSlotSave=async({status,note,moveTo})=>{
    let slots=[...data.slots];
    const idx=slots.findIndex(s=>s.id===activeSlot.id);
    if(moveTo){ const ti=slots.findIndex(s=>s.id===moveTo); slots[ti]={...slots[ti],booking:slots[idx].booking,type:"booked",status:"changed",note:"שונה שעה"}; slots[idx]={...slots[idx],booking:null,type:"free",status:null,note:""}; }
    else slots[idx]={...slots[idx],status,note};
    await save({...data,slots}); setActiveSlot(null);
  };

  const booked=data.slots.filter(s=>s.booking);
  const arrived=data.slots.filter(s=>s.status==="arrived");
  const cancelled=data.slots.filter(s=>s.status==="cancelled");
  const noshow=data.slots.filter(s=>s.status==="noshow");
  const pending=booked.filter(s=>!s.status);
  const free=data.slots.filter(s=>s.type==="free"&&!s.booking);
  const regNames=booked.map(s=>s.booking.childName.trim().toLowerCase());
  const unregistered=(data.students||[]).filter(s=>!regNames.includes(s.name.trim().toLowerCase()));

  const slotStyle={
    arrived:{bg:C.arrived,border:C.arrivedB,tag:{color:C.accent,label:"✅ הגיע/ה"}},
    cancelled:{bg:C.cancelled,border:C.cancelledB,tag:{color:C.danger,label:"❌ ביטל/ה"}},
    noshow:{bg:C.noshow,border:C.noshowB,tag:{color:C.warm,label:"⏳ לא הגיע/ה"}},
    changed:{bg:C.changed,border:C.changedB,tag:{color:"#5B6FD4",label:"🔄 שונה שעה"}},
  };

  return(
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Heebo','Arial Hebrew',Arial,sans-serif",direction:"rtl",color:C.text}}>
      <div style={{background:C.primary,position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 12px #0004"}}>
        <div style={{maxWidth:620,margin:"0 auto",padding:"12px 16px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div>
            <h1 style={{color:"white",margin:0,fontSize:16,fontWeight:800}}>📅 {data.teacherName}{data.subject?` · ${data.subject}`:""}</h1>
            <div style={{color:"#A8D5B5",fontSize:12,marginTop:1}}>{fmtDate(data.date)}</div>
          </div>
          <div style={{display:"flex",gap:7}}>
            <Btn sm color={C.warm} onClick={()=>setShareOpen(true)}>📤 הזמנה</Btn>
            <Btn sm color={C.purple} onClick={()=>setRemindOpen(true)}>🔔</Btn>
            <Btn sm color={C.gold} disabled={exporting} onClick={async()=>{setExporting(true);try{await exportXLSX(data);}catch(e){alert(e.message);}finally{setExporting(false);}}}>
              {exporting?"...":"📊"}
            </Btn>
          </div>
        </div>
      </div>

      <div style={{maxWidth:620,margin:"0 auto",padding:"16px 14px"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:7,marginBottom:16}}>
          {[["נרשמו",booked.length,C.primary],["הגיעו",arrived.length,C.accent],["ביטלו",cancelled.length,C.danger],["ממתין",pending.length,C.warm],["פנויים",free.length,C.muted]].map(([l,v,col])=>(
            <div key={l} style={{background:C.card,border:`1.5px solid ${C.border}`,borderRadius:10,padding:"9px 6px",textAlign:"center"}}>
              <div style={{fontSize:20,fontWeight:900,color:col}}>{v}</div>
              <div style={{fontSize:11,color:C.muted}}>{l}</div>
            </div>
          ))}
        </div>

        {unregistered.length>0&&(
          <div style={{background:C.noshow,border:`1.5px solid ${C.noshowB}55`,borderRadius:12,padding:"12px 16px",marginBottom:14,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div style={{fontSize:14,color:C.warm,fontWeight:700}}>⚠️ {unregistered.length} תלמידים עדיין לא נרשמו</div>
            <Btn sm color={C.warm} onClick={()=>setRemindOpen(true)}>שלח תזכורת</Btn>
          </div>
        )}

        <div style={{display:"flex",gap:0,marginBottom:14,background:C.border+"55",borderRadius:10,padding:3}}>
          {[["schedule","לוח שעות"],["tracking","מעקב ביום"],["roster","רשימת תלמידים"]].map(([t,l])=>(
            <button key={t} onClick={()=>setTab(t)} style={{flex:1,padding:"8px",borderRadius:8,border:"none",background:tab===t?C.card:"transparent",color:tab===t?C.primary:C.muted,fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit",transition:"all .15s"}}>{l}</button>
          ))}
        </div>

        {tab==="schedule"&&(
          <div style={{display:"flex",flexDirection:"column",gap:7}}>
            {data.slots.map(sl=>{
              const isBooked=!!sl.booking,isBreak=sl.type==="break";
              const ss=sl.status?slotStyle[sl.status]:null;
              let bg=C.card,border=C.border;
              if(isBreak){bg="#F3EFF9";border=C.purple;}
              else if(ss){bg=ss.bg;border=ss.border;}
              else if(isBooked){bg="#F0F7FF";border="#90BEE8";}
              return(
                <div key={sl.id} style={{display:"flex",alignItems:"center",gap:10,padding:"11px 14px",borderRadius:10,border:`1.5px solid ${border}`,background:bg,cursor:isBooked?"pointer":"default",transition:"all .15s"}} onClick={()=>isBooked&&setActiveSlot(sl)}>
                  <span style={{fontWeight:700,fontSize:13,direction:"ltr",whiteSpace:"nowrap",minWidth:100}}>{sl.start} – {sl.end}</span>
                  <div style={{flex:1,minWidth:0}}>
                    {isBooked&&<div><span style={{fontWeight:700,fontSize:14,color:C.primary}}>{sl.booking.parentName}</span><span style={{color:C.muted,fontSize:13}}> · {sl.booking.childName}</span></div>}
                    {isBreak&&<span style={{fontSize:13,color:C.purple,fontWeight:600}}>☕ הפסקה</span>}
                    {!isBooked&&!isBreak&&<span style={{fontSize:13,color:C.muted}}>פנוי</span>}
                  </div>
                  <div style={{display:"flex",gap:6,alignItems:"center"}}>
                    {ss&&<Tag color={ss.tag.color}>{ss.tag.label}</Tag>}
                    {sl.note&&<span title={sl.note} style={{fontSize:14,cursor:"help"}}>📝</span>}
                    {!isBooked&&<button onClick={e=>{e.stopPropagation();toggleBreak(sl.id);}} style={{background:isBreak?C.accent+"22":C.purple+"18",color:isBreak?C.accent:C.purple,border:`1px solid ${isBreak?C.accent:C.purple}33`,borderRadius:7,width:28,height:28,cursor:"pointer",fontSize:14,fontFamily:"inherit"}}>{isBreak?"↺":"☕"}</button>}
                    {isBooked&&<button onClick={e=>{e.stopPropagation();cancelSlot(sl.id);}} style={{background:C.danger+"18",color:C.danger,border:`1px solid ${C.danger}33`,borderRadius:7,width:28,height:28,cursor:"pointer",fontWeight:800,fontSize:12,fontFamily:"inherit"}}>✕</button>}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {tab==="tracking"&&(
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            {pending.length>0&&(
              <div>
                <div style={{fontSize:13,fontWeight:700,color:C.warm,marginBottom:8,display:"flex",alignItems:"center",gap:6}}>
                  <span style={{background:C.warm,color:"white",borderRadius:20,padding:"2px 9px",fontSize:12}}>{pending.length}</span> ממתינים לסימון
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:6}}>
                  {pending.map(sl=>(
                    <div key={sl.id} style={{padding:"12px 14px",borderRadius:10,border:`1.5px solid ${C.border}`,background:C.card,display:"flex",alignItems:"center",justifyContent:"space-between",cursor:"pointer"}} onClick={()=>setActiveSlot(sl)}>
                      <div>
                        <div style={{fontWeight:700,color:C.primary}}>{sl.booking.parentName} <span style={{color:C.muted,fontWeight:400}}>· {sl.booking.childName}</span></div>
                        <div style={{fontSize:12,color:C.muted,direction:"ltr",marginTop:2}}>{sl.start} – {sl.end}{sl.booking.phone&&` · ${sl.booking.phone}`}</div>
                      </div>
                      <span style={{color:C.muted,fontSize:13}}>סמן ←</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {[{list:arrived,label:"הגיעו",color:C.accent,icon:"✅"},{list:cancelled,label:"ביטלו",color:C.danger,icon:"❌"},{list:noshow,label:"לא הגיעו",color:C.warm,icon:"⏳"}].filter(g=>g.list.length>0).map(g=>(
              <div key={g.label}>
                <div style={{fontSize:13,fontWeight:700,color:g.color,marginBottom:8,display:"flex",alignItems:"center",gap:6}}>
                  <span style={{background:g.color,color:"white",borderRadius:20,padding:"2px 9px",fontSize:12}}>{g.list.length}</span> {g.label}
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:6}}>
                  {g.list.map(sl=>(
                    <div key={sl.id} style={{padding:"11px 14px",borderRadius:10,border:`1.5px solid ${g.color}44`,background:g.color+"11",display:"flex",alignItems:"center",justifyContent:"space-between",cursor:"pointer"}} onClick={()=>setActiveSlot(sl)}>
                      <div>
                        <div style={{fontWeight:700,color:C.primary}}>{sl.booking.parentName} <span style={{color:C.muted,fontWeight:400}}>· {sl.booking.childName}</span></div>
                        <div style={{fontSize:12,color:C.muted,direction:"ltr",marginTop:1}}>{sl.start}{sl.note&&` · 📝 ${sl.note}`}</div>
                      </div>
                      <span style={{fontSize:18}}>{g.icon}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
            {booked.length===0&&<div style={{textAlign:"center",padding:"40px 20px",color:C.muted}}><div style={{fontSize:40,marginBottom:10}}>📭</div>אין הרשמות עדיין</div>}
          </div>
        )}

        {tab==="roster"&&(
          <div>
            {(!data.students||data.students.length===0)
              ?<div style={{textAlign:"center",padding:"40px 20px",color:C.muted}}><div style={{fontSize:40,marginBottom:10}}>📋</div>לא הוזנה רשימת תלמידים</div>
              :<>
                <div style={{fontSize:13,fontWeight:700,color:C.muted,marginBottom:12}}>{data.students.length} תלמידים · {unregistered.length} לא נרשמו</div>
                <div style={{display:"flex",flexDirection:"column",gap:6}}>
                  {data.students.map(s=>{
                    const isReg=regNames.includes(s.name.trim().toLowerCase());
                    const slot=isReg?data.slots.find(sl=>sl.booking?.childName.trim().toLowerCase()===s.name.trim().toLowerCase()):null;
                    return(
                      <div key={s.id} style={{padding:"10px 14px",borderRadius:10,border:`1.5px solid ${isReg?C.arrivedB+"55":C.noshowB+"44"}`,background:isReg?C.arrived:C.noshow,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                        <div>
                          <div style={{fontWeight:700,fontSize:14,color:C.primary}}>{s.name}</div>
                          {s.phone&&<div style={{fontSize:12,color:C.muted,marginTop:1}}>{s.phone}</div>}
                        </div>
                        {isReg?<Tag color={C.accent}>✓ {slot?.start}</Tag>:<Tag color={C.warm}>⏳ לא נרשם/ה</Tag>}
                      </div>
                    );
                  })}
                </div>
              </>
            }
          </div>
        )}
      </div>

      <Footer/>
      {activeSlot&&<SlotModal slot={activeSlot} allSlots={data.slots} onClose={()=>setActiveSlot(null)} onSave={handleSlotSave}/>}
      {shareOpen&&<ShareModal data={data} sessionId={sessionId} onClose={()=>setShareOpen(false)}/>}
      {remindOpen&&<RemindersModal data={data} sessionId={sessionId} onClose={()=>setRemindOpen(false)}/>}
    </div>
  );
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
export default function App(){
  const params=new URLSearchParams(window.location.search);
  const urlSid=params.get("s");      // parent link
  const urlAdmin=params.get("admin"); // teacher admin link

  const [phase,setPhase]=useState(()=>{
    if(urlSid) return "loading";
    if(urlAdmin) return "loading";
    return "welcome";
  });
  const [data,setData]=useState(null);
  const [sid,setSid]=useState(null);
  const [newData,setNewData]=useState(null); // data just created, waiting for SaveLinks

  useEffect(()=>{
    const load=async()=>{
      if(urlAdmin){
        const d=await store.get(`session:${urlAdmin}`);
        if(d){setData(d);setSid(urlAdmin);setPhase("dashboard");}
        else setPhase("notfound");
      } else if(urlSid){
        const d=await store.get(`session:${urlSid}`);
        if(d){setData(d);setSid(urlSid);setPhase("parent");}
        else setPhase("notfound");
      }
    };
    if(urlSid||urlAdmin) load();
  },[]);

  const handleSetupComplete=d=>{ setNewData(d); setPhase("savelinks"); };
  const handleSaveDone=()=>{ setData(newData); setSid(newData.sid); setPhase("dashboard"); };

  const handleBook=async(slotId,booking)=>{
    const slots=data.slots.map(s=>s.id!==slotId?s:{...s,type:"booked",booking});
    const updated={...data,slots};
    await store.set(`session:${sid}`,updated);
    setData(updated);
  };

  if(phase==="loading") return(
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Heebo','Arial Hebrew',Arial,sans-serif",direction:"rtl"}}>
      <div style={{textAlign:"center",color:C.muted}}><div style={{fontSize:48,marginBottom:12}}>📅</div><div>טוען...</div></div>
    </div>
  );
  if(phase==="notfound") return(
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Heebo','Arial Hebrew',Arial,sans-serif",direction:"rtl"}}>
      <div style={{textAlign:"center"}}><div style={{fontSize:48,marginBottom:12}}>🔍</div><div style={{fontWeight:700,color:C.text,marginBottom:6,fontSize:18}}>קישור לא נמצא</div><div style={{fontSize:14,color:C.muted}}>ייתכן שהמורה טרם יצר את הלוח</div></div>
    </div>
  );
  if(phase==="welcome")   return <Welcome onStart={()=>setPhase("setup")}/>;
  if(phase==="setup")     return <Setup onComplete={handleSetupComplete}/>;
  if(phase==="savelinks") return <SaveLinks data={newData} sessionId={newData.sid} onDone={handleSaveDone}/>;
  if(phase==="dashboard") return <Dashboard data={data} onUpdate={setData} sessionId={sid}/>;
  if(phase==="parent")    return <ParentView data={data} onBook={handleBook}/>;
  return null;
}
