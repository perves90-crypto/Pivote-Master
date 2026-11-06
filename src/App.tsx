import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Row = Record<string, any>;
type Agg = "SUM" | "COUNT" | "COUNTA" | "AVERAGE" | "MIN" | "MAX";
export default function App() {
  const [book,setBook]=useState<XLSX.WorkBook|null>(null);
  const [sheet,setSheet]=useState("");
  const [rows,setRows]=useState<Row[]>([]);
  const [headers,setHeaders]=useState<string[]>([]);
  const [rowField,setRowField]=useState("");
  const [colField,setColField]=useState("");
  const [valueField,setValueField]=useState("");
  const [filterField,setFilterField]=useState("");
  const [filterValue,setFilterValue]=useState("");
  const [agg,setAgg]=useState<Agg>("SUM");
  const [fileName,setFileName]=useState("");
  const [message,setMessage]=useState("Excel ফাইল খুলুন অথবা নমুনা ডেটা দিয়ে শুরু করুন।");
  const [tab,setTab]=useState<"data"|"pivot">("data");
  const [query,setQuery]=useState("");
  const loadSheet=(wb:XLSX.WorkBook,name:string)=>{
    const ws=wb.Sheets[name]; const data=XLSX.utils.sheet_to_json<Row>(ws,{defval:""});
    setSheet(name);setRows(data);setHeaders(data.length?Object.keys(data[0]):[]);
    setRowField("");setColField("");setValueField("");setFilterField("");setFilterValue("");
    setMessage(data.length+"টি রেকর্ড লোড হয়েছে।");
  };
  const openFile=async(e:React.ChangeEvent<HTMLInputElement>)=>{
    const f=e.target.files?.[0];if(!f)return;
    try{const b=await f.arrayBuffer();const wb=XLSX.read(b,{type:"array",cellDates:true});if(!wb.SheetNames.length)throw new Error("কোনো Sheet পাওয়া যায়নি");
      setBook(wb);setFileName(f.name);loadSheet(wb,wb.SheetNames[0]);}
    catch(err){setMessage("ফাইল খোলা যায়নি: "+String(err));}
    e.target.value="";
  };
  const sample=()=>{
    const data=[{বিভাগ:"খাদ্য",পণ্য:"বিস্কুট",এলাকা:"যশোর",বিক্রয়:12000},{বিভাগ:"খাদ্য",পণ্য:"কেক",এলাকা:"খুলনা",বিক্রয়:8500},{বিভাগ:"পানীয়",পণ্য:"জুস",এলাকা:"যশোর",বিক্রয়:6500},{বিভাগ:"খাদ্য",পণ্য:"বিস্কুট",এলাকা:"খুলনা",বিক্রয়:9000},{বিভাগ:"পানীয়",পণ্য:"জুস",এলাকা:"ঢাকা",বিক্রয়:11000}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),"Sales");setBook(wb);setFileName("নমুনা-বিক্রয়.xlsx");loadSheet(wb,"Sales");
  };
  const filtered=useMemo(()=>rows.filter(r=>(!filterField||!filterValue||String(r[filterField]).toLowerCase().includes(filterValue.toLowerCase()))&&(!query||Object.values(r).some(v=>String(v).toLowerCase().includes(query.toLowerCase())))),[rows,filterField,filterValue,query]);
  const pivot=useMemo(()=>{
    if(!rowField||!valueField)return [];
    const groups=new Map<string,number[]>();
    filtered.forEach(r=>{const k=String(r[rowField]??"(খালি)");const v=Number(r[valueField]);const arr=groups.get(k)||[];arr.push(Number.isFinite(v)?v:0);groups.set(k,arr);});
    const calc=(a:number[])=>agg==="COUNT"?a.filter(x=>Number.isFinite(x)).length:agg==="COUNTA"?a.length:agg==="AVERAGE"?(a.length?a.reduce((x,y)=>x+y,0)/a.length:0):agg==="MIN"?Math.min(...a):agg==="MAX"?Math.max(...a):a.reduce((x,y)=>x+y,0);
    return [...groups.entries()].map(([k,a])=>({label:k,value:calc(a),count:a.length})).sort((a,b)=>String(a.label).localeCompare(String(b.label),"bn"));
  },[filtered,rowField,valueField,agg]);
  const total=pivot.reduce((s,r)=>s+r.value,0);
  const exportPivot=()=>{const out=pivot.map(x=>({[rowField]:x.label,[agg+" "+valueField]:x.value,রেকর্ড:x.count}));out.push({[rowField]:"সর্বমোট",[agg+" "+valueField]:total,রেকর্ড:filtered.length} as any);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(out),"Pivot");XLSX.writeFile(wb,"Pivot-Master-Report.xlsx");};
  const exportData=()=>{const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(filtered),"Data");XLSX.writeFile(wb,"Pivot-Master-Data.xlsx");};
  return <main>
    <header><div className="brand"><div className="logo">PM</div><div><h1>Pivot Master</h1><small>Excel Pivot Table • Mobile</small></div></div><span className="offline">● Offline</span></header>
    <section className="hero"><p className="eyebrow">SMART SPREADSHEET TOOL</p><h2>আপনার Excel ডেটা,<br/>এখন হাতের মুঠোয়।</h2><p>ফাইল খুলুন, ডেটা বিশ্লেষণ করুন এবং Pivot Report তৈরি করুন।</p><label className="primary filebtn">＋ Excel ফাইল খুলুন<input type="file" accept=".xlsx,.xls,.csv" onChange={openFile}/></label><button className="secondary" onClick={sample}>নমুনা ডেটা দিয়ে দেখুন</button></section>
    {fileName&&<section className="filebar"><span>📄 {fileName} <small>· {sheet} · {rows.length} rows</small></span><button onClick={()=>{setBook(null);setRows([]);setHeaders([]);setFileName("");}}>বন্ধ করুন</button></section>}
    {headers.length>0&&<><nav><button className={tab==="data"?"active":""} onClick={()=>setTab("data")}>ডেটা দেখুন</button><button className={tab==="pivot"?"active":""} onClick={()=>setTab("pivot")}>Pivot Table</button></nav>
    {book&&book.SheetNames.length>1&&<div className="sheetpick"><label>Sheet নির্বাচন</label><select value={sheet} onChange={e=>loadSheet(book,e.target.value)}>{book.SheetNames.map(s=><option key={s}>{s}</option>)}</select></div>}
    {tab==="data"?<section className="panel"><div className="panelhead"><div><h3>ডেটা প্রিভিউ</h3><p>{filtered.length} / {rows.length} রেকর্ড</p></div><button onClick={exportData}>Export XLSX ↓</button></div><input className="search" placeholder="ডেটার মধ্যে খুঁজুন..." value={query} onChange={e=>setQuery(e.target.value)}/><div className="tablewrap"><table><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{filtered.slice(0,100).map((r,i)=><tr key={i}>{headers.map(h=><td key={h}>{String(r[h])}</td>)}</tr>)}</tbody></table></div><small className="hint">প্রিভিউতে সর্বোচ্চ ১০০টি সারি দেখানো হচ্ছে।</small></section>:
    <section className="panel"><div className="panelhead"><div><h3>Pivot Table তৈরি করুন</h3><p>কোন তথ্য কীভাবে সাজাবেন নির্বাচন করুন</p></div><span className="tag">PIVOT</span></div><div className="fieldgrid"><label>Rows (সারি)<select value={rowField} onChange={e=>setRowField(e.target.value)}><option value="">Field নির্বাচন</option>{headers.map(h=><option key={h}>{h}</option>)}</select></label><label>Columns (কলাম)<select value={colField} onChange={e=>setColField(e.target.value)}><option value="">নেই</option>{headers.filter(h=>h!==rowField).map(h=><option key={h}>{h}</option>)}</select></label><label>Values (মান)<select value={valueField} onChange={e=>setValueField(e.target.value)}><option value="">Field নির্বাচন</option>{headers.map(h=><option key={h}>{h}</option>)}</select></label><label>Calculation<select value={agg} onChange={e=>setAgg(e.target.value as Agg)}>{(["SUM","COUNT","COUNTA","AVERAGE","MIN","MAX"] as Agg[]).map(x=><option key={x}>{x}</option>)}</select></label><label>Filter Field<select value={filterField} onChange={e=>{setFilterField(e.target.value);setFilterValue("");}}><option value="">কোনো Filter নয়</option>{headers.map(h=><option key={h}>{h}</option>)}</select></label><label>Filter Value<input value={filterValue} onChange={e=>setFilterValue(e.target.value)} placeholder="যে মানটি রাখতে চান"/></label></div>
    {rowField&&valueField?<><div className="resulthead"><div><h3>Pivot Result</h3><p>{agg} of {valueField} by {rowField}</p></div><button className="primary smallbtn" onClick={exportPivot}>Excel Export ↓</button></div><div className="tablewrap"><table><thead><tr><th>{rowField}</th><th>{agg} of {valueField}</th><th>রেকর্ড</th></tr></thead><tbody>{pivot.map(r=><tr key={r.label}><td>{r.label}</td><td>{r.value.toLocaleString("en-US",{maximumFractionDigits:2})}</td><td>{r.count}</td></tr>)}<tr className="total"><td>সর্বমোট</td><td>{total.toLocaleString("en-US",{maximumFractionDigits:2})}</td><td>{filtered.length}</td></tr></tbody></table></div></>:<div className="empty">Pivot তৈরি করতে Rows এবং Values নির্বাচন করুন।</div>}</section>}</>}
    <footer>Pivot Master <span>· আপনার ফাইল আপনার ডিভাইসেই থাকে</span></footer><div className="toast">{message}</div>
  </main>;
}