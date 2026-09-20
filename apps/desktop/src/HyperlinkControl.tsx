import {useEffect,useState} from 'react'
import {RibbonButton} from './Ribbon'
export default function HyperlinkControl({url,disabled,onChange}:{url?:string;disabled:boolean;onChange(url:string|null):void}){
 const [open,setOpen]=useState(false),[address,setAddress]=useState(url??'')
 useEffect(()=>{setAddress(url??'');setOpen(false)},[url])
 return <div className="office-link-control"><RibbonButton icon="link" label={url?'Edit link':'Insert link'} disabled={disabled} aria-expanded={open} onClick={()=>{setAddress(url??'');setOpen(value=>!value)}} />{open&&<form aria-label="Link settings" onSubmit={event=>{event.preventDefault();onChange(address.trim());setOpen(false)}}>
  <label>Link address<input aria-label="Link address" placeholder="https://example.com" maxLength={2048} value={address} disabled={disabled} onChange={event=>setAddress(event.target.value)} /></label>
  <small>Applies to the selected text segment. Web and email links are supported.</small>
  <RibbonButton icon="check" label="Apply link" type="submit" disabled={disabled||!address.trim()} />{url&&<RibbonButton icon="deleteObject" label="Remove link" disabled={disabled} onClick={()=>{onChange(null);setOpen(false)}} />}<RibbonButton icon="close" label="Cancel" onClick={()=>setOpen(false)} />
 </form>}</div>
}
