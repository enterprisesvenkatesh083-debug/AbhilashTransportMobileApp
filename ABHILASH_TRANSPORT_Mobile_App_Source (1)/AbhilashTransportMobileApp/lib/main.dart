import 'dart:io';
import 'package:flutter/material.dart';
import 'package:sqflite/sqflite.dart';
import 'package:path/path.dart' as p;
import 'package:excel/excel.dart';
import 'package:share_plus/share_plus.dart';
import 'package:path_provider/path_provider.dart';
import 'package:intl/intl.dart';

void main() async {
 WidgetsFlutterBinding.ensureInitialized();
 await DB.instance.open();
 runApp(const App());
}
class Trip {
 final int? id; final String date,vehicle,fromTime,toTime,journey,user,driver,totalHrs,exHrs; final double fromKm,toKm;
 Trip({this.id,required this.date,required this.vehicle,required this.fromTime,required this.toTime,required this.fromKm,required this.toKm,required this.journey,required this.user,required this.driver,this.totalHrs='',this.exHrs=''});
 double get km=>toKm-fromKm;
 Map<String,Object?> map()=>{'date':date,'vehicle':vehicle,'fromTime':fromTime,'toTime':toTime,'fromKm':fromKm,'toKm':toKm,'totalKm':km,'journey':journey,'user':user,'driver':driver,'totalHrs':totalHrs,'exHrs':exHrs};
 factory Trip.from(Map<String,Object?> m)=>Trip(id:m['id'] as int?,date:'${m['date']}',vehicle:'${m['vehicle']}',fromTime:'${m['fromTime']}',toTime:'${m['toTime']}',fromKm:(m['fromKm'] as num).toDouble(),toKm:(m['toKm'] as num).toDouble(),journey:'${m['journey']}',user:'${m['user']}',driver:'${m['driver']}',totalHrs:'${m['totalHrs']??''}',exHrs:'${m['exHrs']??''}');
}
class DB {
 DB._(); static final instance=DB._(); Database? db;
 Future<void> open() async { final path=p.join(await getDatabasesPath(),'abhilash_logbook.db'); db=await openDatabase(path,version:1,onCreate:(d,v)=>d.execute('''CREATE TABLE trips(id INTEGER PRIMARY KEY AUTOINCREMENT,date TEXT,vehicle TEXT,fromTime TEXT,toTime TEXT,fromKm REAL,toKm REAL,totalKm REAL,journey TEXT,user TEXT,driver TEXT,totalHrs TEXT,exHrs TEXT)''')); }
 Future<List<Trip>> all() async => (await db!.query('trips',orderBy:'date DESC,id DESC')).map(Trip.from).toList();
 Future<void> add(Trip t) async { await db!.insert('trips',t.map()); }
}
class App extends StatelessWidget { const App({super.key}); @override Widget build(BuildContext c)=>MaterialApp(title:'ABHILASH TRANSPORT',debugShowCheckedModeBanner:false,theme:ThemeData(colorScheme:ColorScheme.fromSeed(seedColor:const Color(0xff174a78)),useMaterial3:true),home:const Home()); }
class Home extends StatefulWidget { const Home({super.key}); @override State<Home> createState()=>_HomeState(); }
class _HomeState extends State<Home> {
 List<Trip> trips=[]; bool busy=false;
 @override void initState(){super.initState();refresh();}
 Future<void> refresh() async { final x=await DB.instance.all();if(mounted)setState(()=>trips=x); }
 void note(String s)=>ScaffoldMessenger.of(context).showSnackBar(SnackBar(content:Text(s)));
 Future<void> shareExcel() async {
  if(trips.isEmpty){note('पहले New Entry से रिकॉर्ड सेव करें।');return;}
  setState(()=>busy=true);
  try {
   final excel=Excel.createExcel(); final sheet=excel['Vehicle Log'];
   final heads=['SR NO','DATE','VEHICLE NUMBER','FROM TIME','TO TIME','FROM KM','TO KM','TOTAL KM','BRIEF DETAILS OF JOURNEY','NAME OF USER','TOTAL HRS','EX. HRS','DRIVER'];
   sheet.appendRow(heads.map((e)=>TextCellValue(e)).toList());
   for(int i=0;i<trips.length;i++){final t=trips[i];sheet.appendRow([IntCellValue(i+1),TextCellValue(t.date),TextCellValue(t.vehicle),TextCellValue(t.fromTime),TextCellValue(t.toTime),DoubleCellValue(t.fromKm),DoubleCellValue(t.toKm),DoubleCellValue(t.km),TextCellValue(t.journey),TextCellValue(t.user),TextCellValue(t.totalHrs),TextCellValue(t.exHrs),TextCellValue(t.driver)]);}
   for(int i=0;i<heads.length;i++){sheet.setColumnWidth(i,[9,14,20,12,12,12,12,12,38,22,12,12,20][i].toDouble());sheet.cell(CellIndex.indexByColumnRow(columnIndex:i,rowIndex:0)).cellStyle=CellStyle(bold:true,fontColorHex:ExcelColor.white,backgroundColorHex:ExcelColor.fromHexString('#174A78'));}
   final bytes=excel.encode();if(bytes==null)throw Exception('Excel create failed');
   final dir=await getTemporaryDirectory();final name='ABHILASH_TRANSPORT_Logbook_${DateFormat('yyyyMMdd_HHmm').format(DateTime.now())}.xlsx';final file=File(p.join(dir.path,name));await file.writeAsBytes(bytes,flush:true);
   await Share.shareXFiles([XFile(file.path,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',name:name)],subject:'ABHILASH TRANSPORT Logbook',text:'Vehicle logbook Excel report');
  }catch(e){note('Excel share error: $e');}finally{if(mounted)setState(()=>busy=false);}
 }
 @override Widget build(BuildContext c)=>Scaffold(appBar:AppBar(title:const Text('ABHILASH TRANSPORT'),backgroundColor:const Color(0xffdff0f5)),floatingActionButton:FloatingActionButton.extended(onPressed:()async{final ok=await Navigator.push<bool>(c,MaterialPageRoute(builder:(_)=>const Entry()));if(ok==true)refresh();},icon:const Icon(Icons.add),label:const Text('New Entry')),body:RefreshIndicator(onRefresh:refresh,child:ListView(padding:const EdgeInsets.all(12),children:[
  Card(child:Padding(padding:const EdgeInsets.all(16),child:Column(crossAxisAlignment:CrossAxisAlignment.start,children:[
   const Text('Daily Staff Vehicle Log Book',style:TextStyle(fontWeight:FontWeight.bold,fontSize:19)),Text('${trips.length} saved entries'),
   const SizedBox(height:12),SizedBox(width:double.infinity,child:FilledButton.icon(onPressed:busy?null:shareExcel,icon:const Icon(Icons.share),label:Text(busy?'Preparing Excel…':'Export & Share Excel'))),
   const Text('Share menu se WhatsApp, Email, Drive ya dusri app par Excel bhejein.',style:TextStyle(fontSize:12))
  ]))),
  if(trips.isEmpty)const Padding(padding:EdgeInsets.all(30),child:Center(child:Text('अभी कोई एंट्री नहीं है। New Entry दबाएँ।')))
  else ...trips.map((t)=>Card(child:ListTile(title:Text('${t.date} • ${t.vehicle}',style:const TextStyle(fontWeight:FontWeight.bold)),subtitle:Text('${t.fromTime}–${t.toTime} | ${t.fromKm} → ${t.toKm} KM\n${t.journey}\nUser: ${t.user} • Driver: ${t.driver}'),isThreeLine:true,trailing:Text('${t.km.toStringAsFixed(0)} KM')))),
  const SizedBox(height:80)
 ])));
}
class Entry extends StatefulWidget {const Entry({super.key});@override State<Entry> createState()=>_EntryState();}
class _EntryState extends State<Entry>{
 final form=GlobalKey<FormState>();
 late final date=TextEditingController(text:DateFormat('yyyy-MM-dd').format(DateTime.now()));
 final vehicle=TextEditingController(text:'MH12ZU6903'),fromTime=TextEditingController(),toTime=TextEditingController(),fromKm=TextEditingController(),toKm=TextEditingController(),journey=TextEditingController(),user=TextEditingController(),driver=TextEditingController(),totalHrs=TextEditingController(),exHrs=TextEditingController();
 Duration? get durationWorked { final a=_parseClock(fromTime.text), b=_parseClock(toTime.text); if(a==null||b==null)return null; var mins=b-a; if(mins<0)mins+=24*60; return Duration(minutes:mins); }
 int? _parseClock(String s){ final m=RegExp(r'^(\\d{1,2}):(\\d{2})
 bool saving=false; String? req(String? v)=>v==null||v.trim().isEmpty?'Required':null;
 double? get km{final a=double.tryParse(fromKm.text),b=double.tryParse(toKm.text);return a!=null&&b!=null&&b>=a?b-a:null;}
 Widget field(String label,TextEditingController c,{bool required=true,TextInputType? type})=>TextFormField(controller:c,keyboardType:type,decoration:InputDecoration(labelText:label),validator:required?req:null,onChanged:(_)=>setState((){if(identical(c,fromTime)||identical(c,toTime))_updateHours();}));
 Future<void> chooseDate()async{final d=await showDatePicker(context:context,initialDate:DateTime.tryParse(date.text)??DateTime.now(),firstDate:DateTime(2020),lastDate:DateTime(2100));if(d!=null)setState(()=>date.text=DateFormat('yyyy-MM-dd').format(d));}
 Future<void> chooseTime(TextEditingController c)async{final t=await showTimePicker(context:context,initialTime:TimeOfDay.now());if(t!=null)setState((){c.text='${t.hour.toString().padLeft(2,'0')}:${t.minute.toString().padLeft(2,'0')}';_updateHours();});}
 Future<void> save()async{if(!form.currentState!.validate())return;final a=double.tryParse(fromKm.text),b=double.tryParse(toKm.text);if(a==null||b==null||b<a){ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content:Text('Ending KM cannot be less than Starting KM')));return;}setState(()=>saving=true);try{await DB.instance.add(Trip(date:date.text,vehicle:vehicle.text.trim().toUpperCase(),fromTime:fromTime.text,toTime:toTime.text,fromKm:a,toKm:b,journey:journey.text.trim(),user:user.text.trim(),driver:driver.text.trim(),totalHrs:totalHrs.text.trim(),exHrs:exHrs.text.trim()));if(mounted)Navigator.pop(context,true);}catch(e){if(mounted)ScaffoldMessenger.of(context).showSnackBar(SnackBar(content:Text('Save failed: $e')));}finally{if(mounted)setState(()=>saving=false);}}
 @override void dispose(){for(final x in [date,vehicle,fromTime,toTime,fromKm,toKm,journey,user,driver,totalHrs,exHrs]){x.dispose();}super.dispose();}
 @override Widget build(BuildContext c)=>Scaffold(appBar:AppBar(title:const Text('New Journey Entry')),body:Form(key:form,child:ListView(padding:const EdgeInsets.all(16),children:[
  InkWell(onTap:chooseDate,child:InputDecorator(decoration:const InputDecoration(labelText:'Date *',suffixIcon:Icon(Icons.calendar_month)),child:Text(date.text))),const SizedBox(height:12),
  field('Vehicle Number *',vehicle),const SizedBox(height:12),
  Row(children:[Expanded(child:InkWell(onTap:()=>chooseTime(fromTime),child:InputDecorator(decoration:const InputDecoration(labelText:'From Time *'),child:Text(fromTime.text.isEmpty?'Select time':fromTime.text)))),const SizedBox(width:10),Expanded(child:InkWell(onTap:()=>chooseTime(toTime),child:InputDecorator(decoration:const InputDecoration(labelText:'To Time *'),child:Text(toTime.text.isEmpty?'Select time':toTime.text))))]),const SizedBox(height:12),
  Row(children:[Expanded(child:field('Starting KM *',fromKm,type:const TextInputType.numberWithOptions(decimal:true))),const SizedBox(width:10),Expanded(child:field('Ending KM *',toKm,type:const TextInputType.numberWithOptions(decimal:true)))]),
  const SizedBox(height:8),Container(padding:const EdgeInsets.all(12),color:Theme.of(c).colorScheme.surfaceContainerHighest,child:Text('Total KM: ${km?.toStringAsFixed(1)??'—'}',style:const TextStyle(fontWeight:FontWeight.bold))),const SizedBox(height:12),
  field('Brief Details of Journey *',journey),const SizedBox(height:12),field('Name of User *',user),const SizedBox(height:12),field('Driver Name *',driver),const SizedBox(height:12),
  Row(children:[Expanded(child:InputDecorator(decoration:const InputDecoration(labelText:'TOTAL HRS (auto)'),child:Text(totalHrs.text.isEmpty?'—':totalHrs.text))),const SizedBox(width:10),Expanded(child:InputDecorator(decoration:const InputDecoration(labelText:'EX. HRS (after 16 hrs)'),child:Text(exHrs.text.isEmpty?'—':exHrs.text)))]),const SizedBox(height:20),
  FilledButton.icon(onPressed:saving?null:save,icon:const Icon(Icons.save),label:Text(saving?'Saving…':'Save Entry'))
 ])));
}
).firstMatch(s.trim()); if(m==null)return null; final h=int.tryParse(m.group(1)!), n=int.tryParse(m.group(2)!); if(h==null||n==null||h>23||n>59)return null; return h*60+n; }
 String _formatDuration(Duration d){ final mins=d.inMinutes; return '${mins~/60}:${(mins%60).toString().padLeft(2,'0')}'; }
 void _updateHours(){ final d=durationWorked; if(d==null){totalHrs.clear();exHrs.clear();}else{totalHrs.text=_formatDuration(d); final extra=d-Duration(hours:16); exHrs.text=_formatDuration(extra.isNegative?Duration.zero:extra);} }
 bool saving=false; String? req(String? v)=>v==null||v.trim().isEmpty?'Required':null;
 double? get km{final a=double.tryParse(fromKm.text),b=double.tryParse(toKm.text);return a!=null&&b!=null&&b>=a?b-a:null;}
 Widget field(String label,TextEditingController c,{bool required=true,TextInputType? type})=>TextFormField(controller:c,keyboardType:type,decoration:InputDecoration(labelText:label),validator:required?req:null,onChanged:(_)=>setState((){}));
 Future<void> chooseDate()async{final d=await showDatePicker(context:context,initialDate:DateTime.tryParse(date.text)??DateTime.now(),firstDate:DateTime(2020),lastDate:DateTime(2100));if(d!=null)setState(()=>date.text=DateFormat('yyyy-MM-dd').format(d));}
 Future<void> chooseTime(TextEditingController c)async{final t=await showTimePicker(context:context,initialTime:TimeOfDay.now());if(t!=null)setState(()=>c.text='${t.hour.toString().padLeft(2,'0')}:${t.minute.toString().padLeft(2,'0')}');}
 Future<void> save()async{if(!form.currentState!.validate())return;final a=double.tryParse(fromKm.text),b=double.tryParse(toKm.text);if(a==null||b==null||b<a){ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content:Text('Ending KM cannot be less than Starting KM')));return;}setState(()=>saving=true);try{await DB.instance.add(Trip(date:date.text,vehicle:vehicle.text.trim().toUpperCase(),fromTime:fromTime.text,toTime:toTime.text,fromKm:a,toKm:b,journey:journey.text.trim(),user:user.text.trim(),driver:driver.text.trim(),totalHrs:totalHrs.text.trim(),exHrs:exHrs.text.trim()));if(mounted)Navigator.pop(context,true);}catch(e){if(mounted)ScaffoldMessenger.of(context).showSnackBar(SnackBar(content:Text('Save failed: $e')));}finally{if(mounted)setState(()=>saving=false);}}
 @override void dispose(){for(final x in [date,vehicle,fromTime,toTime,fromKm,toKm,journey,user,driver,totalHrs,exHrs]){x.dispose();}super.dispose();}
 @override Widget build(BuildContext c)=>Scaffold(appBar:AppBar(title:const Text('New Journey Entry')),body:Form(key:form,child:ListView(padding:const EdgeInsets.all(16),children:[
  InkWell(onTap:chooseDate,child:InputDecorator(decoration:const InputDecoration(labelText:'Date *',suffixIcon:Icon(Icons.calendar_month)),child:Text(date.text))),const SizedBox(height:12),
  field('Vehicle Number *',vehicle),const SizedBox(height:12),
  Row(children:[Expanded(child:InkWell(onTap:()=>chooseTime(fromTime),child:InputDecorator(decoration:const InputDecoration(labelText:'From Time *'),child:Text(fromTime.text.isEmpty?'Select time':fromTime.text)))),const SizedBox(width:10),Expanded(child:InkWell(onTap:()=>chooseTime(toTime),child:InputDecorator(decoration:const InputDecoration(labelText:'To Time *'),child:Text(toTime.text.isEmpty?'Select time':toTime.text))))]),const SizedBox(height:12),
  Row(children:[Expanded(child:field('Starting KM *',fromKm,type:const TextInputType.numberWithOptions(decimal:true))),const SizedBox(width:10),Expanded(child:field('Ending KM *',toKm,type:const TextInputType.numberWithOptions(decimal:true)))]),
  const SizedBox(height:8),Container(padding:const EdgeInsets.all(12),color:Theme.of(c).colorScheme.surfaceContainerHighest,child:Text('Total KM: ${km?.toStringAsFixed(1)??'—'}',style:const TextStyle(fontWeight:FontWeight.bold))),const SizedBox(height:12),
  field('Brief Details of Journey *',journey),const SizedBox(height:12),field('Name of User *',user),const SizedBox(height:12),field('Driver Name *',driver),const SizedBox(height:12),
  Row(children:[Expanded(child:field('Total Hours',totalHrs,required:false)),const SizedBox(width:10),Expanded(child:field('Extra Hours',exHrs,required:false))]),const SizedBox(height:20),
  FilledButton.icon(onPressed:saving?null:save,icon:const Icon(Icons.save),label:Text(saving?'Saving…':'Save Entry'))
 ])));
}
