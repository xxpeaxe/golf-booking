import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const COURSES   = [{id:"valley",name:"Valley",emoji:"🌄"},{id:"lake",name:"Lake",emoji:"🌊"},{id:"hill",name:"Hill",emoji:"⛰️"},{id:"any",name:"상관없음",emoji:"🎯"}];
const TIMES     = ["06:00","06:30","07:00","07:30","08:00","08:30","09:00","09:30","10:00","10:30","11:00","11:30","12:00","13:00","14:00","15:00"];
const DAYS      = ["일","월","화","수","목","금","토"];
const OWNER     = "xxpeaxe@naver.com";
const ADMIN_PW  = "fila3470";
const WD_LIMIT  = 10;
const WE_LIMIT  = 5;
const G = "#16a34a", LG = "#f0fdf4";

// ── EmailJS 설정 ──
const EJS_SERVICE_ID       = "service_epsa8vb";
const EJS_PUBLIC_KEY       = "Yh-UG4eCJF5r2USyR";
const EJS_TEMPLATE_CONFIRM = "template_dyjb04r"; // 예약 확정
const EJS_TEMPLATE_GENERIC = "template_21mjuoo"; // 공용 (취소 확인 + 관리자 알림)
const ADMIN_EMAILS         = "mjseo@mistobrand.com,kenneth.shin@mistobrand.com"; // 신청 알림 수신

const addMins  = (t,m)=>{ const [h,min]=t.split(":").map(Number),tot=h*60+min+m; return `${String(Math.floor(tot/60)).padStart(2,"0")}:${String(tot%60).padStart(2,"0")}`; };
const slotLabel= t=>`${t} - ${addMins(t,30)}`;
const isWE     = d=>d.getDay()===0||d.getDay()===6;
const ym       = (y,m)=>`${y}-${String(m+1).padStart(2,"0")}`;
const dStr     = d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const fmt      = d=>d?`${d.getFullYear()}. ${d.getMonth()+1}. ${d.getDate()}. (${DAYS[d.getDay()]})`:"";
const fmtS     = d=>`${d.getMonth()+1}/${d.getDate()}(${DAYS[d.getDay()]})`;

const getNth = (y,mo,dow,n)=>{ let c=0; for(let day=1;day<=31;day++){const dt=new Date(y,mo,day);if(dt.getMonth()!==mo)break;if(dt.getDay()===dow){c++;if(c===n)return dt;}} return null; };
const getWdWindow = (y,mo)=>{ let pm=mo-1,py=y; if(pm<0){pm=11;py--;} const wed=getNth(py,pm,3,1); if(!wed)return{open:null,close:null}; return{open:null,close:new Date(wed.getFullYear(),wed.getMonth(),wed.getDate(),17,0)}; };
const isBookable = (d,now)=>{ if(d<=now)return false; if(isWE(d)){ const ref=new Date(d);ref.setDate(d.getDate()-21); const dow=ref.getDay(),diff=dow===0?-6:1-dow,mon=new Date(ref);mon.setDate(ref.getDate()+diff); return now<=new Date(mon.getFullYear(),mon.getMonth(),mon.getDate(),17,0); } const{close}=getWdWindow(d.getFullYear(),d.getMonth()); return!!(close&&now<=close); };

// ── localStorage 헬퍼 ──
const sGet  = k=>{ try{const v=localStorage.getItem(k);return v?JSON.parse(v):null;}catch{return null;} };
const sSet  = (k,v)=>{ try{localStorage.setItem(k,JSON.stringify(v));return true;}catch{return false;} };
const sList = p=>{ try{return Object.keys(localStorage).filter(k=>k.startsWith(p));}catch{return[];} };
const sDel  = k=>{ try{localStorage.removeItem(k);return true;}catch{return false;} };

// ── EmailJS 발송 ──
const ejsSend = async (templateId, params) => {
  const res = await fetch("https://api.emailjs.com/api/v1.0/email/send",{
    method:"POST", headers:{"Content-Type":"application/json"},
    body:JSON.stringify({service_id:EJS_SERVICE_ID,template_id:templateId,user_id:EJS_PUBLIC_KEY,template_params:params})
  });
  if(!res.ok) throw new Error(`EmailJS ${res.status}`);
};

function ConfirmModal({message,onOk,onCancel}){
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200}}>
      <div style={{background:"#fff",borderRadius:14,padding:"28px 24px",width:280,textAlign:"center",boxShadow:"0 8px 32px rgba(0,0,0,0.18)"}}>
        <p style={{margin:"0 0 24px",fontSize:15,color:"#111",lineHeight:1.6}}>{message}</p>
        <div style={{display:"flex",gap:10}}>
          <button onClick={onCancel} style={{flex:1,padding:"10px",borderRadius:8,border:"1.5px solid #e5e7eb",background:"#fff",cursor:"pointer",fontWeight:600,color:"#374151"}}>취소</button>
          <button onClick={onOk} style={{flex:1,padding:"10px",borderRadius:8,border:"none",background:"#ef4444",color:"#fff",cursor:"pointer",fontWeight:700}}>확인</button>
        </div>
      </div>
    </div>
  );
}

export default function App(){
  const [mode,setMode]=useState("booking");
  if(mode==="admin")return <AdminPanel onExit={()=>setMode("booking")}/>;
  return <BookingPage onAdmin={()=>setMode("admin")}/>;
}

// ══════════════════════════════════════
//  예약 신청 페이지
// ══════════════════════════════════════
function BookingPage({onAdmin}){
  const [step,setStep]           = useState(1);
  const [selDate,setSelDate]     = useState(null);
  const [selTime,setSelTime]     = useState(null);
  const [selCourse,setSelCourse] = useState(null);
  const [players,setPlayers]     = useState(2);
  const [name,setName]           = useState("");
  const [phone,setPhone]         = useState("");
  const [email,setEmail]         = useState("");
  const [kakao,setKakao]         = useState("");
  const [done,setDone]           = useState(false);
  const [bookedTimes,setBookedTimes]   = useState([]);
  const [bookedDates,setBookedDates]   = useState(new Set());
  const [monthlyCount,setMonthlyCount] = useState({wd:0,we:0});
  const [quotaErr,setQuotaErr]   = useState("");
  const [checking,setChecking]   = useState(false);

  const now   = new Date();
  const today = new Date(now.getFullYear(),now.getMonth(),now.getDate());
  const [calY,setCalY]=useState(now.getFullYear());
  const [calM,setCalM]=useState(now.getMonth());

  const wdWin  = getWdWindow(calY,calM);
  const wdInfo = wdWin.close?`${wdWin.close.getMonth()+1}월 ${wdWin.close.getDate()}일 17:00까지`:"-";

  useEffect(()=>{
    const keys=sList(`bookings:${ym(calY,calM)}-`);
    const dates=keys.map(k=>k.split(":")[1]);
    setBookedDates(new Set(dates));
    let wd=0,we=0; dates.forEach(ds=>isWE(new Date(ds))?we++:wd++);
    setMonthlyCount({wd,we});
  },[calY,calM]);

  useEffect(()=>{
    if(step===2&&selDate){
      const prefix=`bookings:${dStr(selDate)}:`;
      const keys=sList(prefix);
      setBookedTimes(keys.map(k=>k.replace(prefix,"").replace("-",":")));
    }
  },[step,selDate]);

  const daysInMonth=new Date(calY,calM+1,0).getDate();
  const firstDay=new Date(calY,calM,1).getDay();
  const prevM=()=>{ const nm=calM===0?11:calM-1,ny=calM===0?calY-1:calY; if(ny<today.getFullYear()||(ny===today.getFullYear()&&nm<today.getMonth()))return; setCalM(nm);setCalY(ny); };
  const nextM=()=>calM===11?(setCalY(y=>y+1),setCalM(0)):setCalM(m=>m+1);
  const canNext=()=>[!!selDate,!!selTime,!!selCourse,!!(name.trim()&&phone.trim()&&email.trim()&&kakao),true][step-1];

  const checkQuota=()=>{
    const mems=sGet("admin:members")||[];
    const mem=mems.find(m=>m.name.trim()===name.trim()&&m.phone.replace(/\D/g,"")===phone.replace(/\D/g,""));
    if(!mem)return "등록된 멤버가 아닙니다. 관리자에게 문의해 주세요.";
    const keys=sList(`bookings:${ym(selDate.getFullYear(),selDate.getMonth())}-`);
    let wd=0,we=0;
    for(const key of keys){const d=sGet(key);if(d?.name===mem.name&&d?.phone===mem.phone){isWE(new Date(key.split(":")[1]))?we++:wd++;}}
    if(isWE(selDate)&&we>=mem.weekendQ)return `주말 예약 한도(${mem.weekendQ}회) 초과 — 현재 ${we}회 사용`;
    if(!isWE(selDate)&&wd>=mem.weekdayQ)return `평일 예약 한도(${mem.weekdayQ}회) 초과 — 현재 ${wd}회 사용`;
    return null;
  };

  const handleNext=()=>{
    if(step===4){setChecking(true);setQuotaErr("");const err=checkQuota();setChecking(false);if(err){setQuotaErr(err);return;}}
    step<5?setStep(s=>s+1):handleConfirm();
  };

  const reset=()=>{setStep(1);setSelDate(null);setSelTime(null);setSelCourse(null);setPlayers(2);setName("");setPhone("");setEmail("");setKakao("");setDone(false);setBookedTimes([]);setQuotaErr("");};

  const handleConfirm=()=>{
    const key=`bookings:${dStr(selDate)}:${selTime.replace(":","-")}`;
    if(sGet(key)){alert("이미 예약된 시간입니다.");setStep(2);return;}
    sSet(key,{name,phone,email,kakao,course:selCourse?.name,players,bookedAt:new Date().toISOString()});
    // 관리자 2명에게 신청 알림 자동 발송 (공용 템플릿)
    const adminMsg=`새로운 골프 예약 신청이 들어왔습니다.\n\n▸ 예약자: ${name}\n▸ 연락처: ${phone}\n▸ 이메일: ${email}\n▸ 카카오톡 알림: ${kakao}\n▸ 날짜: ${fmt(selDate)}\n▸ 티타임: ${selTime} - ${addMins(selTime,30)}\n▸ 코스: ${selCourse?.name}\n▸ 인원: ${players}명\n\n관리자 페이지에서 확인 후 예약을 진행해 주세요.`;
    ejsSend(EJS_TEMPLATE_GENERIC,{to_email:ADMIN_EMAILS,subject:`[골프 예약 신청] ${name} ${fmt(selDate)} ${selTime}`,message:adminMsg}).catch(e=>console.error("관리자 알림 발송 실패:",e));
    setDone(true);
  };

  const inp=(label,val,set,ph,type="text")=>(
    <div key={label}>
      <label style={{fontSize:13,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>{label}</label>
      <input type={type} value={val} onChange={e=>set(e.target.value)} placeholder={ph}
        style={{width:"100%",padding:"11px 13px",borderRadius:8,border:"1.5px solid #e5e7eb",fontSize:14,outline:"none",boxSizing:"border-box"}}
        onFocus={e=>e.target.style.borderColor=G} onBlur={e=>e.target.style.borderColor="#e5e7eb"}/>
    </div>
  );

  const stepLabels=["날짜 선택","티타임 선택","코스 & 인원","정보 입력","최종 확인"];

  if(done)return(
    <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",justifyContent:"center",alignItems:"center",padding:16}}>
      <div style={{background:"#fff",borderRadius:20,boxShadow:"0 2px 12px rgba(0,0,0,0.08)",width:"100%",maxWidth:460,padding:"40px 28px",textAlign:"center"}}>
        <div style={{fontSize:64}}>✅</div>
        <h2 style={{color:G,margin:"16px 0 8px",fontSize:22}}>신청이 완료됐어요!</h2>
        <div style={{background:LG,borderRadius:12,padding:"18px 20px",textAlign:"left",margin:"20px 0",border:"1px solid #bbf7d0"}}>
          {[["날짜",fmt(selDate)],["티타임",`${selTime} - ${addMins(selTime,30)}`],["코스",selCourse?.name],["인원",`${players}명`],["예약자",name],["연락처",phone],["카카오 알림",kakao==="O"?"필요":"불필요"]].map(([k,v])=>(
            <div key={k} style={{display:"flex",justifyContent:"space-between",marginBottom:8,fontSize:14}}>
              <span style={{color:"#6b7280"}}>{k}</span><span style={{fontWeight:600,color:"#111"}}>{v}</span>
            </div>
          ))}
        </div>
        <button onClick={reset} style={{background:G,color:"#fff",border:"none",borderRadius:10,padding:"13px 32px",cursor:"pointer",fontWeight:700,fontSize:15}}>새 예약 신청하기</button>
      </div>
    </div>
  );

  return(
    <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",justifyContent:"center",alignItems:"flex-start",padding:"32px 16px",fontFamily:"'Inter','Apple SD Gothic Neo',sans-serif"}}>
      <div style={{background:"#fff",borderRadius:20,boxShadow:"0 2px 12px rgba(0,0,0,0.08)",width:"100%",maxWidth:460,padding:"32px 28px"}}>
        <div style={{marginBottom:28}}>
          <h1 style={{margin:0,fontSize:20,fontWeight:800,color:"#111"}}>⛳ 골프 예약 신청</h1>
          <p style={{margin:"6px 0 0",fontSize:12,color:"#9ca3af"}}>※ 예약 신청으로 예약 확정 아님, 확정 시 개별 안내</p>
          <div style={{display:"flex",gap:5,marginTop:12}}>{[1,2,3,4,5].map(i=><div key={i} style={{flex:1,height:4,borderRadius:2,background:i<=step?G:"#e5e7eb",transition:"background 0.3s"}}/>)}</div>
          <p style={{margin:"8px 0 0",color:"#9ca3af",fontSize:13}}>Step {step} / 5 — {stepLabels[step-1]}</p>
          <div style={{marginTop:10,padding:"9px 12px",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,fontSize:12,color:"#92400e",lineHeight:1.6}}>
            ℹ️ 잔여 횟수가 초과되거나 희망 일자가 겹칠 경우, <strong>선착순</strong>으로 예약 신청이 진행됩니다.
          </div>
        </div>

        {step===1&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <button onClick={prevM} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:"#374151"}}>‹</button>
              <span style={{fontWeight:700,color:"#111",fontSize:16}}>{calY}년 {calM+1}월</span>
              <button onClick={nextM} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:"#374151"}}>›</button>
            </div>
            <div style={{background:"#f9fafb",borderRadius:8,padding:"9px 12px",marginBottom:12,fontSize:11,color:"#6b7280",lineHeight:1.8}}>
              <div>📋 <strong>평일</strong> 신청 마감: {wdInfo}</div>
              <div>🏌️ <strong>주말</strong> 신청 마감: 희망일 기준 3주 전 월요일 17:00까지</div>
              <div style={{marginTop:4,paddingTop:4,borderTop:"1px solid #e5e7eb",display:"flex",gap:14}}>
                <span>📊 이번 달 잔여</span>
                <span style={{fontWeight:600,color:monthlyCount.wd>=WD_LIMIT?"#ef4444":G}}>평일 {WD_LIMIT-monthlyCount.wd}/{WD_LIMIT}</span>
                <span style={{fontWeight:600,color:monthlyCount.we>=WE_LIMIT?"#ef4444":"#3b82f6"}}>주말 {WE_LIMIT-monthlyCount.we}/{WE_LIMIT}</span>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4,marginBottom:6}}>
              {DAYS.map(d=><div key={d} style={{textAlign:"center",fontSize:11,fontWeight:600,padding:"4px 0",color:d==="일"?"#ef4444":d==="토"?"#3b82f6":"#9ca3af"}}>{d}</div>)}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>
              {Array(firstDay).fill(null).map((_,i)=><div key={`e${i}`}/>)}
              {Array(daysInMonth).fill(null).map((_,i)=>{
                const day=i+1,d=new Date(calY,calM,day),past=d<today,bookable=!past&&isBookable(d,now);
                const sel=selDate&&selDate.getTime()===d.getTime(),dow=d.getDay();
                const fullBooked=bookedDates.has(dStr(d)),limitReached=!fullBooked&&(isWE(d)?monthlyCount.we>=WE_LIMIT:monthlyCount.wd>=WD_LIMIT);
                const blocked=fullBooked||limitReached,disabled=!bookable||blocked;
                return(
                  <button key={day} disabled={disabled} onClick={()=>setSelDate(d)} style={{aspectRatio:"1",borderRadius:8,border:"none",cursor:disabled?"not-allowed":"pointer",background:sel?G:blocked?"#fff1f2":"transparent",color:past?"#e5e7eb":blocked?"#fca5a5":!bookable?"#d1d5db":sel?"#fff":dow===0?"#ef4444":dow===6?"#3b82f6":"#111",fontWeight:sel?700:400,fontSize:12,lineHeight:1.2,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
                    <span>{day}</span>{blocked&&<span style={{fontSize:7,color:"#f87171",fontWeight:700}}>마감</span>}
                  </button>
                );
              })}
            </div>
            {selDate&&<p style={{textAlign:"center",color:G,fontWeight:700,marginTop:16,fontSize:14}}>✓ {fmt(selDate)}</p>}
          </div>
        )}

        {step===2&&(
          <div>
            <p style={{color:"#6b7280",fontSize:13,marginBottom:14}}>{fmt(selDate)} 티타임을 선택해 주세요.</p>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8}}>
              {TIMES.map(t=>{const booked=bookedTimes.includes(t),sel=selTime===t;return(
                <button key={t} disabled={booked} onClick={()=>setSelTime(t)} style={{padding:"11px 10px",borderRadius:8,border:`1.5px solid ${sel?G:booked?"#f3f4f6":"#e5e7eb"}`,background:sel?LG:booked?"#f9fafb":"#fff",color:sel?G:booked?"#d1d5db":"#374151",fontWeight:sel?700:400,cursor:booked?"not-allowed":"pointer",fontSize:13}}>
                  {slotLabel(t)}{booked&&<div style={{fontSize:9,color:"#d1d5db"}}>예약됨</div>}
                </button>
              );})}
            </div>
            {bookedTimes.length>0&&<p style={{fontSize:12,color:"#9ca3af",marginTop:12,textAlign:"center"}}>회색 슬롯은 이미 예약된 시간입니다.</p>}
          </div>
        )}

        {step===3&&(
          <div>
            <p style={{color:"#6b7280",fontSize:13,marginBottom:10}}>코스를 선택해 주세요.</p>
            <div style={{display:"flex",flexDirection:"column",gap:9,marginBottom:24}}>
              {COURSES.map(c=>(
                <button key={c.id} onClick={()=>setSelCourse(c)} style={{display:"flex",alignItems:"center",gap:14,padding:"14px 16px",borderRadius:10,border:`1.5px solid ${selCourse?.id===c.id?G:"#e5e7eb"}`,background:selCourse?.id===c.id?LG:"#fff",cursor:"pointer",textAlign:"left"}}>
                  <span style={{fontSize:24}}>{c.emoji}</span>
                  <span style={{fontWeight:700,color:"#111",fontSize:16}}>{c.name}</span>
                  {selCourse?.id===c.id&&<span style={{marginLeft:"auto",color:G,fontSize:18}}>✓</span>}
                </button>
              ))}
            </div>
            <p style={{color:"#6b7280",fontSize:13,marginBottom:10}}>인원 (1~4명)</p>
            <div style={{display:"flex",alignItems:"center",gap:14}}>
              {["-","+"].map((op,idx)=>(
                <button key={op} onClick={()=>setPlayers(p=>idx===0?Math.max(1,p-1):Math.min(4,p+1))} style={{width:38,height:38,borderRadius:8,border:"1.5px solid #e5e7eb",background:"#fff",fontSize:20,cursor:"pointer",fontWeight:700}}>{op}</button>
              ))}
              <span style={{fontSize:22,fontWeight:800,minWidth:28,textAlign:"center"}}>{players}</span>
              <span style={{color:"#9ca3af",fontSize:14}}>명</span>
            </div>
          </div>
        )}

        {step===4&&(
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            {inp("예약자 이름",name,setName,"홍길동")}
            {inp("연락처",phone,setPhone,"010-0000-0000","tel")}
            {inp("이메일",email,setEmail,"example@email.com","email")}
            <div>
              <label style={{fontSize:13,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>카카오톡 알림 발송 필요 여부</label>
              <div style={{display:"flex",gap:8}}>
                {[{v:"O",label:"⭕ 필요",ac:G,bg:LG},{v:"X",label:"❌ 불필요",ac:"#ef4444",bg:"#fff0f0"}].map(({v,label,ac,bg})=>(
                  <button key={v} type="button" onClick={()=>setKakao(v)} style={{flex:1,padding:"11px",borderRadius:8,border:`1.5px solid ${kakao===v?ac:"#e5e7eb"}`,background:kakao===v?bg:"#fff",color:kakao===v?ac:"#374151",fontWeight:kakao===v?700:400,cursor:"pointer",fontSize:14}}>{label}</button>
                ))}
              </div>
            </div>
            {quotaErr&&<div style={{background:"#fff0f0",border:"1px solid #fecaca",borderRadius:8,padding:"10px 13px",fontSize:13,color:"#dc2626"}}>⚠️ {quotaErr}</div>}
          </div>
        )}

        {step===5&&(
          <div>
            <div style={{background:LG,borderRadius:14,padding:"20px",marginBottom:14,border:"1px solid #bbf7d0"}}>
              {[["📅 날짜",fmt(selDate)],["⏰ 티타임",`${selTime} - ${addMins(selTime,30)}`],["⛳ 코스",selCourse?.name],["👥 인원",`${players}명`],["👤 예약자",name],["📞 연락처",phone],["💬 카카오 알림",kakao==="O"?"필요":"불필요"]].map(([k,v])=>(
                <div key={k} style={{display:"flex",justifyContent:"space-between",marginBottom:10,fontSize:14}}>
                  <span style={{color:"#6b7280"}}>{k}</span><span style={{fontWeight:600,color:"#111"}}>{v}</span>
                </div>
              ))}
            </div>
            <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:10,padding:"12px 14px",fontSize:12,color:"#92400e",lineHeight:1.6}}>
              ⚠️ 희망하신 시간에 예약이 어려울 경우, 가장 유사한 시간으로 신청됩니다.
            </div>
          </div>
        )}

        <div style={{display:"flex",gap:10,marginTop:28}}>
          {step>1&&<button onClick={()=>{setQuotaErr("");setStep(s=>s-1);}} style={{flex:1,padding:"12px 0",borderRadius:10,border:"1.5px solid #e5e7eb",background:"#fff",cursor:"pointer",fontWeight:600,color:"#374151",fontSize:15}}>이전</button>}
          <button disabled={!canNext()||checking} onClick={handleNext} style={{flex:2,padding:"12px 0",borderRadius:10,border:"none",background:canNext()&&!checking?G:"#d1d5db",color:"#fff",cursor:canNext()&&!checking?"pointer":"not-allowed",fontWeight:700,fontSize:15}}>
            {checking?"확인 중...":step===5?"신청 완료":"다음 →"}
          </button>
        </div>
        <p onClick={onAdmin} style={{textAlign:"center",marginTop:20,fontSize:11,color:"#d1d5db",cursor:"pointer",userSelect:"none"}}>관리자</p>
      </div>
    </div>
  );
}

// ══════════════════════════════════════
//  관리자 패널
// ══════════════════════════════════════
function AdminPanel({onExit}){
  const [authed,setAuthed]       = useState(false);
  const [pw,setPw]               = useState("");
  const [pwErr,setPwErr]         = useState(false);
  const [tab,setTab]             = useState("members");
  const [members,setMembers]     = useState([]);
  const [bookings,setBookings]   = useState([]);
  const [confirmedKeys,setConfirmedKeys]   = useState(new Set());
  const [kakaoSentKeys,setKakaoSentKeys]   = useState(new Set());
  const [quotaData,setQuotaData] = useState([]);
  const [cancelItems,setCancelItems] = useState([]);
  const [confirm,setConfirm]     = useState(null);
  const [sending,setSending]     = useState(null);

  const now=new Date();
  const todayMid=new Date(now.getFullYear(),now.getMonth(),now.getDate());
  const [fY,setFY]=useState(now.getFullYear());
  const [fM,setFM]=useState(now.getMonth());

  const [nName,setNName]=useState("");
  const [nPhone,setNPhone]=useState("");
  const [nEmail,setNEmail]=useState("");
  const [nWd,setNWd]=useState(10);
  const [nWe,setNWe]=useState(5);
  const [editId,setEditId]=useState(null);

  const login=()=>pw===ADMIN_PW?(setAuthed(true),setPwErr(false)):setPwErr(true);
  const loadMembers=useCallback(()=>{setMembers(sGet("admin:members")||[]);},[]);

  const loadBookings=useCallback(()=>{
    const keys=sList(`bookings:${ym(fY,fM)}-`);
    const res=[]; const cfSet=new Set(); const kkSet=new Set();
    for(const key of keys){
      const d=sGet(key);
      if(d){
        const p=key.split(":");
        res.push({key,ds:p[1],ts:p[2]?.replace("-",":"),...d});
        if(sGet(key.replace("bookings:","confirm:")))cfSet.add(key);
        if(sGet(key.replace("bookings:","kakaosent:")))kkSet.add(key);
      }
    }
    res.sort((a,b)=>a.ds.localeCompare(b.ds)||a.ts.localeCompare(b.ts));
    setBookings(res); setConfirmedKeys(cfSet); setKakaoSentKeys(kkSet);
  },[fY,fM]);

  const loadQuota=useCallback(()=>{
    const mems=sGet("admin:members")||[];
    const keys=sList(`bookings:${ym(fY,fM)}-`);
    const bkgs=[];
    for(const key of keys){const d=sGet(key);if(d){const dt=new Date(key.split(":")[1]);bkgs.push({...d,we:isWE(dt)});}}
    setQuotaData(mems.map(m=>{const mine=bkgs.filter(b=>b.name===m.name&&b.phone===m.phone);return{...m,wdU:mine.filter(b=>!b.we).length,weU:mine.filter(b=>b.we).length};}));
  },[fY,fM]);

  const loadCancel=useCallback(()=>{
    const seen=new Set(),items=[];
    for(let w=0;w<8;w++){
      const ref=new Date(todayMid);ref.setDate(ref.getDate()+w*7);
      const keys=sList(`bookings:${ym(ref.getFullYear(),ref.getMonth())}-`);
      for(const key of keys){
        if(seen.has(key))continue;seen.add(key);
        const d=sGet(key);if(!d)continue;
        const ds=key.split(":")[1],ts=key.split(":")[2]?.replace("-",":");
        const date=new Date(ds);if(date<=todayMid)continue;
        const dow=date.getDay()===0?7:date.getDay();
        const bkgMon=new Date(date);bkgMon.setDate(date.getDate()-dow+1);
        const noticeWed=new Date(bkgMon);noticeWed.setDate(bkgMon.getDate()-7+2); // 전주 수요일
        items.push({key,ds,ts,date,noticeFri:noticeWed,isToday:noticeWed.toDateString()===todayMid.toDateString(),isPast:noticeWed<todayMid,...d});
      }
    }
    items.sort((a,b)=>a.date-b.date);setCancelItems(items);
  },[]);

  useEffect(()=>{if(authed)loadMembers();},[authed]);
  useEffect(()=>{if(!authed)return;if(tab==="bookings")loadBookings();if(tab==="quota")loadQuota();if(tab==="cancel")loadCancel();},[authed,tab,fY,fM]);

  const sendConfirmEmail=async b=>{
    try{
      await ejsSend(EJS_TEMPLATE_CONFIRM,{to_email:b.email,name:b.name,date:b.ds,time:b.ts,course:b.course||"-",players:b.players,reply_email:OWNER});
      return true;
    }catch(e){console.error("EmailJS 오류:",e);return false;}
  };

  const toggleConfirm=async b=>{
    const cfKey=b.key.replace("bookings:","confirm:");
    if(confirmedKeys.has(b.key)){
      sDel(cfKey);
      setConfirmedKeys(prev=>{const s=new Set(prev);s.delete(b.key);return s;});
    }else{
      setSending(b.key);
      const ok=await sendConfirmEmail(b);
      setSending(null);
      if(!ok){alert("메일 발송 실패. EmailJS 설정을 확인해 주세요.");return;}
      sSet(cfKey,true);
      setConfirmedKeys(prev=>new Set([...prev,b.key]));
    }
  };

  const toggleKakaoSent=b=>{
    const kkKey=b.key.replace("bookings:","kakaosent:");
    if(kakaoSentKeys.has(b.key)){sDel(kkKey);setKakaoSentKeys(prev=>{const s=new Set(prev);s.delete(b.key);return s;});}
    else{sSet(kkKey,true);setKakaoSentKeys(prev=>new Set([...prev,b.key]));}
  };

  const saveMember=()=>{
    if(!nName.trim()||!nPhone.trim())return;
    const mems=sGet("admin:members")||[];
    if(editId){sSet("admin:members",mems.map(m=>m.id===editId?{...m,name:nName.trim(),phone:nPhone.trim(),email:nEmail.trim(),weekdayQ:nWd,weekendQ:nWe}:m));setEditId(null);}
    else{sSet("admin:members",[...mems,{id:Date.now().toString(),name:nName.trim(),phone:nPhone.trim(),email:nEmail.trim(),weekdayQ:nWd,weekendQ:nWe}]);}
    setNName("");setNPhone("");setNEmail("");setNWd(10);setNWe(5);loadMembers();
  };
  const startEdit=m=>{setEditId(m.id);setNName(m.name);setNPhone(m.phone);setNEmail(m.email||"");setNWd(m.weekdayQ);setNWe(m.weekendQ);};
  const cancelEdit=()=>{setEditId(null);setNName("");setNPhone("");setNEmail("");setNWd(10);setNWe(5);};
  const deleteMember=id=>setConfirm({message:"멤버를 삭제할까요?",onOk:()=>{const mems=sGet("admin:members")||[];sSet("admin:members",mems.filter(m=>m.id!==id));loadMembers();}});
  const deleteBooking=key=>setConfirm({message:"예약을 삭제할까요?",onOk:()=>{sDel(key);loadBookings();}});

  // 취소 여부 확인 메일 (공용 템플릿)
  const sendCancelMail=async item=>{
    const mems=sGet("admin:members")||[];
    const toEmail=mems.find(m=>m.name===item.name&&m.phone===item.phone)?.email||item.email;
    if(!toEmail){alert("이메일 정보가 없습니다.");return;}
    const msg=`안녕하세요, ${item.name}님.\n\n다가오는 골프 예약 건의 취소 여부를 확인드립니다.\n\n▸ 예약일: ${item.ds}\n▸ 티타임: ${item.ts}\n▸ 코스: ${item.course||"-"}\n▸ 인원: ${item.players}명\n\n취소를 원하시면 이 메일에 회신 부탁드립니다.\n별도 회신이 없으실 경우 예약은 그대로 유지됩니다.\n\n문의: ${OWNER}\n\n감사합니다.`;
    try{
      await ejsSend(EJS_TEMPLATE_GENERIC,{to_email:toEmail,subject:`[골프 예약 취소 여부 확인] ${item.name}님 ${item.ds}`,message:msg});
      alert(`${item.name}님에게 취소 확인 메일을 발송했습니다.`);
    }catch(e){alert("메일 발송 실패. EmailJS 설정을 확인해 주세요.");}
  };

  const exportExcel=()=>{
    if(!bookings.length){alert("내보낼 예약이 없습니다.");return;}
    const rows=bookings.map(b=>{const date=new Date(b.ds);return{
      "날짜":b.ds,"요일":DAYS[date.getDay()],"구분":isWE(date)?"주말":"평일",
      "티타임":b.ts,"예약자":b.name,"연락처":b.phone,"이메일":b.email||"-",
      "코스":b.course||"-","인원":b.players,
      "예약 확정":confirmedKeys.has(b.key)?"✓":"",
      "카카오 알림":b.kakao==="O"?"필요":"불필요",
      "카카오 발송":b.kakao==="O"?(kakaoSentKeys.has(b.key)?"✓":"대기"):"-",
      "신청일시":b.bookedAt?new Date(b.bookedAt).toLocaleString("ko-KR"):"-",
    };});
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=[{wch:12},{wch:6},{wch:6},{wch:12},{wch:10},{wch:14},{wch:22},{wch:8},{wch:5},{wch:10},{wch:10},{wch:10},{wch:20}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,`${fY}년 ${fM+1}월`);
    XLSX.writeFile(wb,`골프예약_${ym(fY,fM)}.xlsx`);
  };

  const prevM=()=>fM===0?(setFY(y=>y-1),setFM(11)):setFM(m=>m-1);
  const nextM=()=>fM===11?(setFY(y=>y+1),setFM(0)):setFM(m=>m+1);

  const inp2=(label,val,set,ph,type="text")=>(
    <div><label style={{fontSize:11,fontWeight:600,color:"#6b7280",display:"block",marginBottom:4}}>{label}</label>
    <input type={type} value={val} onChange={e=>set(e.target.value)} placeholder={ph} style={{width:"100%",padding:"8px 10px",borderRadius:6,border:"1.5px solid #e5e7eb",fontSize:13,outline:"none",boxSizing:"border-box"}} onFocus={e=>e.target.style.borderColor=G} onBlur={e=>e.target.style.borderColor="#e5e7eb"}/></div>
  );
  const numInp=(label,val,set)=>(
    <div><label style={{fontSize:11,fontWeight:600,color:"#6b7280",display:"block",marginBottom:4}}>{label}</label>
    <input type="number" value={val} onChange={e=>set(Number(e.target.value))} min={0} max={50} style={{width:"100%",padding:"8px 10px",borderRadius:6,border:"1.5px solid #e5e7eb",fontSize:13,outline:"none",boxSizing:"border-box"}}/></div>
  );
  const bar=(used,total,color)=>(
    <div style={{height:6,background:"#f3f4f6",borderRadius:3,marginTop:4}}>
      <div style={{height:"100%",borderRadius:3,width:`${Math.min(100,(used/total)*100)}%`,background:used>=total?"#ef4444":color}}/>
    </div>
  );

  if(!authed)return(
    <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",justifyContent:"center",alignItems:"center",fontFamily:"'Inter','Apple SD Gothic Neo',sans-serif"}}>
      <div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 12px rgba(0,0,0,0.08)",padding:"40px 36px",width:300,textAlign:"center"}}>
        <div style={{fontSize:44,marginBottom:8}}>🔒</div>
        <h2 style={{margin:"0 0 24px",fontSize:18,fontWeight:800}}>관리자 로그인</h2>
        <input type="password" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} placeholder="비밀번호"
          style={{width:"100%",padding:"11px 13px",borderRadius:8,border:`1.5px solid ${pwErr?"#ef4444":"#e5e7eb"}`,fontSize:14,outline:"none",boxSizing:"border-box",marginBottom:8}}/>
        {pwErr&&<p style={{color:"#ef4444",fontSize:12,margin:"0 0 10px"}}>비밀번호가 틀렸습니다.</p>}
        <button onClick={login} style={{width:"100%",padding:"12px",borderRadius:8,border:"none",background:G,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:15}}>로그인</button>
        <button onClick={onExit} style={{marginTop:12,background:"none",border:"none",color:"#9ca3af",cursor:"pointer",fontSize:12}}>← 예약 페이지로</button>
      </div>
    </div>
  );

  const TABS=[{id:"members",l:"👥 멤버"},{id:"bookings",l:"📅 예약"},{id:"quota",l:"📊 쿼터"},{id:"cancel",l:"🔔 취소알림"}];
  const todayAlerts=cancelItems.filter(i=>i.isToday);

  return(
    <div style={{minHeight:"100vh",background:"#f9fafb",fontFamily:"'Inter','Apple SD Gothic Neo',sans-serif"}}>
      {confirm&&<ConfirmModal message={confirm.message} onOk={()=>{confirm.onOk();setConfirm(null);}} onCancel={()=>setConfirm(null)}/>}
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontWeight:800,fontSize:16,color:"#111"}}>⛳ 골프 예약 관리자</span>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {todayAlerts.length>0&&<span style={{background:"#f97316",color:"#fff",borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:700}}>🔔 오늘 {todayAlerts.length}건</span>}
          <button onClick={()=>{setAuthed(false);onExit();}} style={{background:"none",border:"1px solid #e5e7eb",borderRadius:6,padding:"6px 12px",cursor:"pointer",fontSize:12,color:"#6b7280"}}>나가기</button>
        </div>
      </div>
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",display:"flex",padding:"0 20px"}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"12px 16px",border:"none",background:"none",cursor:"pointer",fontSize:13,fontWeight:tab===t.id?700:400,color:tab===t.id?G:"#6b7280",borderBottom:`2px solid ${tab===t.id?G:"transparent"}`}}>
            {t.l}{t.id==="cancel"&&todayAlerts.length>0&&<span style={{marginLeft:4,background:"#f97316",color:"#fff",borderRadius:10,padding:"1px 6px",fontSize:10}}>{todayAlerts.length}</span>}
          </button>
        ))}
      </div>

      <div style={{padding:"20px",maxWidth:720,margin:"0 auto"}}>

        {tab==="members"&&(
          <div>
            <div style={{background:"#fff",borderRadius:12,padding:"20px",marginBottom:16,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
              <h3 style={{margin:"0 0 14px",fontSize:14,fontWeight:700}}>{editId?"✏️ 멤버 수정":"➕ 멤버 추가"}</h3>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                {inp2("이름",nName,setNName,"홍길동")}
                {inp2("연락처",nPhone,setNPhone,"010-0000-0000","tel")}
                {inp2("이메일",nEmail,setNEmail,"example@email.com","email")}
                <div/>
                {numInp("평일 쿼터 (월)",nWd,setNWd)}
                {numInp("주말 쿼터 (월)",nWe,setNWe)}
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={saveMember} style={{background:G,color:"#fff",border:"none",borderRadius:8,padding:"9px 20px",cursor:"pointer",fontWeight:700,fontSize:13}}>{editId?"수정 완료":"추가"}</button>
                {editId&&<button onClick={cancelEdit} style={{background:"#fff",color:"#6b7280",border:"1px solid #e5e7eb",borderRadius:8,padding:"9px 20px",cursor:"pointer",fontSize:13}}>취소</button>}
              </div>
            </div>
            <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
              <div style={{padding:"14px 20px",borderBottom:"1px solid #f3f4f6",fontWeight:700,fontSize:14}}>전체 멤버 ({members.length}명)</div>
              {members.length===0&&<p style={{textAlign:"center",color:"#9ca3af",padding:24,fontSize:13}}>등록된 멤버가 없습니다.</p>}
              {members.map(m=>(
                <div key={m.id} style={{padding:"13px 20px",borderBottom:"1px solid #f9fafb",display:"flex",alignItems:"center",gap:10}}>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:600,fontSize:14}}>{m.name}</div>
                    <div style={{fontSize:12,color:"#9ca3af"}}>{m.phone} · {m.email||"이메일 없음"}</div>
                    <div style={{fontSize:11,color:"#6b7280",marginTop:2}}>평일 {m.weekdayQ}회 · 주말 {m.weekendQ}회 / 월</div>
                  </div>
                  <button onClick={()=>startEdit(m)} style={{background:"#f9fafb",border:"1px solid #e5e7eb",borderRadius:6,padding:"5px 10px",cursor:"pointer",fontSize:12}}>수정</button>
                  <button onClick={()=>deleteMember(m.id)} style={{background:"#fff0f0",border:"1px solid #fecaca",borderRadius:6,padding:"5px 10px",cursor:"pointer",fontSize:12,color:"#ef4444"}}>삭제</button>
                </div>
              ))}
            </div>
          </div>
        )}

        {tab==="bookings"&&(
          <div>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
              <button onClick={prevM} style={{background:"none",border:"none",cursor:"pointer",fontSize:20}}>‹</button>
              <span style={{fontWeight:700,fontSize:16}}>{fY}년 {fM+1}월</span>
              <button onClick={nextM} style={{background:"none",border:"none",cursor:"pointer",fontSize:20}}>›</button>
              <button onClick={loadBookings} style={{marginLeft:"auto",background:"#f9fafb",border:"1px solid #e5e7eb",borderRadius:6,padding:"6px 12px",cursor:"pointer",fontSize:12}}>새로고침</button>
              <button onClick={exportExcel} style={{background:G,color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",cursor:"pointer",fontSize:12,fontWeight:700}}>📥 엑셀</button>
            </div>
            <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
              <div style={{padding:"13px 20px",borderBottom:"1px solid #f3f4f6",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontWeight:700,fontSize:14}}>예약 {bookings.length}건</span>
                <span style={{fontSize:11,color:"#9ca3af"}}>예약확정 체크 시 확정 메일 자동 발송</span>
              </div>
              {bookings.length===0&&<p style={{textAlign:"center",color:"#9ca3af",padding:24,fontSize:13}}>예약이 없습니다.</p>}
              {bookings.map(b=>{
                const we=isWE(new Date(b.ds));
                const confirmed=confirmedKeys.has(b.key);
                const kakaoSent=kakaoSentKeys.has(b.key);
                const needsKakao=b.kakao==="O";
                const isSending=sending===b.key;
                return(
                  <div key={b.key} style={{padding:"12px 20px",borderBottom:"1px solid #f9fafb",background:confirmed?"#f0fdf4":"#fff",transition:"background 0.2s"}}>
                    <div style={{display:"flex",gap:10,alignItems:"center"}}>
                      <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:2,flexShrink:0}}>
                        {isSending
                          ? <div style={{width:18,height:18,border:"2px solid #d1d5db",borderTopColor:G,borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
                          : <input type="checkbox" checked={confirmed} onChange={()=>toggleConfirm(b)} style={{width:18,height:18,cursor:"pointer",accentColor:G}}/>
                        }
                        <span style={{fontSize:9,color:"#9ca3af",whiteSpace:"nowrap"}}>예약확정</span>
                      </div>
                      <div style={{width:50,textAlign:"center",flexShrink:0}}>
                        <div style={{fontSize:12,fontWeight:600}}>{b.ds.slice(5)}</div>
                        <div style={{fontSize:10,fontWeight:600,color:we?"#3b82f6":G,background:we?"#eff6ff":LG,borderRadius:4,padding:"2px 4px",marginTop:3}}>{we?"주말":"평일"}</div>
                      </div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontWeight:600,fontSize:13,color:confirmed?"#15803d":"#111",display:"flex",alignItems:"center",flexWrap:"wrap",gap:4}}>
                          {b.name} · {b.players}명
                          {confirmed&&<span style={{fontSize:10,color:"#16a34a",background:"#dcfce7",borderRadius:4,padding:"1px 5px"}}>✓ 예약확정</span>}
                          {needsKakao&&<span style={{fontSize:10,background:"#FEE500",color:"#3C1E1E",borderRadius:4,padding:"1px 5px",fontWeight:700}}>카카오</span>}
                        </div>
                        <div style={{fontSize:12,color:"#9ca3af"}}>{b.ts} · {b.course||"-"} · {b.phone}</div>
                      </div>
                      {needsKakao&&(
                        <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:2,flexShrink:0}}>
                          <input type="checkbox" checked={kakaoSent} onChange={()=>toggleKakaoSent(b)} style={{width:18,height:18,cursor:"pointer",accentColor:"#F9E000"}}/>
                          <span style={{fontSize:9,color:"#9ca3af",whiteSpace:"nowrap"}}>카카오발송</span>
                        </div>
                      )}
                      <button onClick={()=>deleteBooking(b.key)} style={{background:"#fff0f0",border:"1px solid #fecaca",borderRadius:6,padding:"4px 8px",cursor:"pointer",fontSize:11,color:"#ef4444",flexShrink:0}}>삭제</button>
                    </div>
                  </div>
                );
              })}
            </div>
            <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
          </div>
        )}

        {tab==="quota"&&(
          <div>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
              <button onClick={prevM} style={{background:"none",border:"none",cursor:"pointer",fontSize:20}}>‹</button>
              <span style={{fontWeight:700,fontSize:16}}>{fY}년 {fM+1}월 쿼터 현황</span>
              <button onClick={nextM} style={{background:"none",border:"none",cursor:"pointer",fontSize:20}}>›</button>
            </div>
            <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
              {quotaData.length===0&&<p style={{textAlign:"center",color:"#9ca3af",padding:24,fontSize:13}}>멤버가 없습니다.</p>}
              {quotaData.map(m=>(
                <div key={m.id} style={{padding:"16px 20px",borderBottom:"1px solid #f9fafb"}}>
                  <div style={{fontWeight:600,fontSize:14,marginBottom:10}}>{m.name}</div>
                  <div style={{marginBottom:8}}>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#6b7280"}}><span>평일</span><span style={{fontWeight:600,color:m.wdU>=m.weekdayQ?"#ef4444":G}}>{m.wdU}/{m.weekdayQ}회</span></div>
                    {bar(m.wdU,m.weekdayQ,G)}
                  </div>
                  <div>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#6b7280"}}><span>주말</span><span style={{fontWeight:600,color:m.weU>=m.weekendQ?"#ef4444":"#3b82f6"}}>{m.weU}/{m.weekendQ}회</span></div>
                    {bar(m.weU,m.weekendQ,"#3b82f6")}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {tab==="cancel"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <span style={{fontWeight:700,fontSize:16}}>취소 알림 발송 일정 (전주 수요일)</span>
              <button onClick={loadCancel} style={{background:"#f9fafb",border:"1px solid #e5e7eb",borderRadius:6,padding:"6px 12px",cursor:"pointer",fontSize:12}}>새로고침</button>
            </div>
            {todayAlerts.length>0&&(
              <div style={{background:"#fff7ed",border:"1px solid #fed7aa",borderRadius:10,padding:"14px 16px",marginBottom:16}}>
                <div style={{fontWeight:700,color:"#92400e",fontSize:13,marginBottom:10}}>🔴 오늘 발송해야 할 알림</div>
                {todayAlerts.map(i=>(
                  <div key={i.key} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderTop:"1px solid #fde68a"}}>
                    <div><div style={{fontWeight:600,fontSize:13}}>{i.name} — {i.ds} {i.ts}</div><div style={{fontSize:11,color:"#92400e"}}>{i.phone}</div></div>
                    <button onClick={()=>sendCancelMail(i)} style={{background:"#f97316",color:"#fff",border:"none",borderRadius:6,padding:"7px 14px",cursor:"pointer",fontSize:12,fontWeight:700}}>📧 발송</button>
                  </div>
                ))}
              </div>
            )}
            <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
              <div style={{padding:"13px 20px",borderBottom:"1px solid #f3f4f6",fontWeight:700,fontSize:14}}>전체 일정 ({cancelItems.length}건)</div>
              {cancelItems.length===0&&<p style={{textAlign:"center",color:"#9ca3af",padding:24,fontSize:13}}>예약이 없습니다.</p>}
              {cancelItems.map(i=>(
                <div key={i.key} style={{padding:"12px 20px",borderBottom:"1px solid #f9fafb",display:"flex",gap:12,alignItems:"center",background:i.isToday?"#fffbeb":i.isPast?"#fafafa":"#fff"}}>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:600,fontSize:13,color:i.isPast?"#9ca3af":"#111"}}>
                      {i.name} — {i.ds} {i.ts}
                      <span style={{marginLeft:6,fontSize:10,fontWeight:600,color:isWE(i.date)?"#3b82f6":G,background:isWE(i.date)?"#eff6ff":LG,borderRadius:4,padding:"1px 5px"}}>{isWE(i.date)?"주말":"평일"}</span>
                    </div>
                    <div style={{fontSize:11,color:"#9ca3af",marginTop:2}}>📧 발송일: {fmtS(i.noticeFri)}{i.isToday?" 🔴 오늘":i.isPast?" (지남)":""}</div>
                  </div>
                  {!i.isPast&&<button onClick={()=>sendCancelMail(i)} style={{background:i.isToday?"#f97316":"#f9fafb",color:i.isToday?"#fff":"#374151",border:`1px solid ${i.isToday?"#f97316":"#e5e7eb"}`,borderRadius:6,padding:"5px 10px",cursor:"pointer",fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>📧 발송</button>}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}