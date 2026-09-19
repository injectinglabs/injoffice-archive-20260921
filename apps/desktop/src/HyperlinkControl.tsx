import {useEffect,useState} from 'react'
export default function HyperlinkControl({url,disabled,onChange}:{url?:string;disabled:boolean;onChange(url:string|null):void}){
 const [open,setOpen]=useState(false),[address,setAddress]=useState(url??'')
 useEffect(()=>{setAddress(url??'');setOpen(false)},[url])
 return <div className="office-link-control"><button disabled={disabled} aria-expanded={open} onClick={()=>{setAddress(url??'');setOpen(value=>!value)}}>{url?'Edit link':'Insert link'}</button>{open&&<form aria-label="Link settings" onSubmit={event=>{event.preventDefault();onChange(address.trim());setOpen(false)}}>
  <label>Link address<input aria-label="Link address" placeholder="https://example.com" maxLength={2048} value={address} disabled={disabled} onChange={event=>setAddress(event.target.value)} /></label>
  <small>Applies to the selected text segment. Web and email links are supported.</small>
  <button type="submit" disabled={disabled||!address.trim()}>Apply link</button>{url&&<button type="button" disabled={disabled} onClick={()=>{onChange(null);setOpen(false)}}>Remove link</button>}<button type="button" onClick={()=>setOpen(false)}>Cancel</button>
 </form>}</div>
}
