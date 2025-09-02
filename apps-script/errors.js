function logError(where, err, extra){
try{
var s = sh_(CONFIG.TABS && CONFIG.TABS.LOGS || 'Logs');
s.appendRow([new Date(),'ERROR',where,String(err && err.stack || err), JSON.stringify(extra||{})]);
}catch(e){ /* swallow */ }
}
function logInfo_(tag, msg){
try{ sh_(CONFIG.TABS && CONFIG.TABS.LOGS || 'Logs').appendRow([new Date(),'INFO',tag,msg||'']); }catch(e){}
}