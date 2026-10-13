import 'dart:io';
import 'package:flutter/material.dart';
import 'package:sqflite/sqflite.dart';
import 'package:path/path.dart' as p;
import 'package:excel/excel.dart';
import 'package:share_plus/share_plus.dart';
import 'package:path_provider/path_provider.dart';
import 'package:intl/intl.dart';

const navy=Color(0xff083b82), blue=Color(0xff087de0), orange=Color(0xffff7900);
void main() async { WidgetsFlutterBinding.ensureInitialized(); await DB.instance.open(); runApp(const App()); }

class Trip {
 final int? id; final String date,vehicle,fromTime,toTime,journey,user,driver,totalHrs,exHrs; final double fromKm,toKm;
 Trip({this.id,required this.date,required this.vehicle,required this.fromTime,required this.toTime,required this.fromKm,required this.toKm,required this.journey,required this.user,required this.driver,required this.totalHrs,required this.exHrs});
 double get km=>toKm-fromKm;
 Map<String,Object?> toMap()=>{'date':date,'vehicle':vehicle,'fromTime':fromTime,'toTime':toTime,'fromKm':fromKm,'toKm':toKm,'totalKm':km,'journey':journey,'user':user,'driver':driver,'totalHrs':totalHrs,'exHrs':exHrs};
 factory Trip.fromMap(Map<String,Object?> m)=>Trip(id:m['id'] as int?,date:'${m['date']}',vehicle:'${m['vehicle']}',fromTime:'${m['fromTime']}',toTime:'${m['toTime']}',fromKm:(m['fromKm'] as num).toDouble(),toKm:(m['toKm'] as num).toDouble(),journey:'${m['journey']}',user:'${m['user']}',driver:'${m['driver']}',totalHrs:'${m['totalHrs']??'0:00'}',exHrs:'${m['exHrs']??'0:00'}');
}
class DB {
 DB._(); static final instance=DB._(); Database? db;
 Future<void> open() async {final path=p.join(await getDatabasesPath(),'abhilash_logbook.db');db=await openDatabase(path,version:1,onCreate:(d,v)=>d.execute('''CREATE TABLE trips(id INTEGER PRIMARY KEY AUTOINCREMENT,date TEXT,vehicle TEXT,fromTime TEXT,toTime TEXT,fromKm REAL,toKm REAL,totalKm REAL,journey TEXT,user TEXT,driver TEXT,totalHrs TEXT,exHrs TEXT)'''));}
 Future<List<Trip>> all() async=>(await db!.query('trips',orderBy:'date DESC,id DESC')).map(Trip.fromMap).toList();
 Future<void> add(Trip t) async {await db!.insert('trips',t.toMap());}
}
class App extends StatelessWidget {const App({super.key});@override Widget build(BuildContext c)=>MaterialApp(title:'ABHILASH TRANSPORT',debugShowCheckedModeBanner:false,theme:ThemeData(colorScheme:ColorScheme.fromSeed(seedColor:blue),useMaterial3:true),home:const Home());}
class BrandLogo extends StatelessWidget {final double height;const BrandLogo({super.key,this.height=60});@override Widget build(BuildContext c)=>SizedBox(height:height,child:Row(mainAxisAlignment:MainAxisAlignment.center,children:[Stack(alignment:Alignment.center,children:[Text('A',style:TextStyle(fontSize:height*.8,fontWeight:FontWeight.w900,color:blue)),Transform.rotate(angle:-.3,child:Container(width:height*.8,height:height*.08,decoration:BoxDecoration(gradient:const LinearGradient(colors:[orange,Color(0xffffc400),Color(0xffff3d00)]),borderRadius:BorderRadius.circular(10))))]),const SizedBox(width:6),Flexible(child:Column(mainAxisAlignment:MainAxisAlignment.center,crossAxisAlignment:CrossAxisAlignment.start,children:[FittedBox(child:Text('ABHILASH',style:TextStyle(fontSize:height*.3,fontWeight:FontWeight.w900,color:navy))),FittedBox(child:Text('TRANSPORT',style:TextStyle(fontSize:height*.2,fontWeight:FontWeight.w900,letterSpacing:2)))]))]));}

class Home extends StatefulWidget {const Home({super.key});@override State<Home> createState()=>_HomeState();}
class _HomeState extends State<Home>{
 List<Trip> trips=[];bool busy=false;
 @override void initState(){super.initState();refresh();}
 Future<void> refresh() async {final x=await DB.instance.all();if(mounted)setState(()=>trips=x);}
 void note(String s)=>ScaffoldMessenger.of(context).showSnackBar(SnackBar(content:Text(s)));
 double get totalKm=>trips.fold<double>(0,(s,t)=>s+t.km);
 int mins(String s){final m=RegExp(r'^(\d+):(\d{2})$').firstMatch(s);return m==null?0:int.parse(m.group(1)!)*60+int.parse(m.group(2)!);}
 String sumHours(Iterable<String> values){final n=values.fold<int>(0,(s,v)=>s+mins(v));return '${n~/60}:${(n%60).toString().padLeft(2,'0')}';}
 Future<void> shareExcel() async {
  if(trips.isEmpty){note('Please add a trip entry first.');return;}setState(()=>busy=true);
  try{
   final book=Excel.createExcel();final dash=book['Dashboard'];
   dash.appendRow([TextCellValue('ABHILASH TRANSPORT')]);dash.appendRow([TextCellValue('VEHICLE LOGBOOK DASHBOARD')]);
   dash.appendRow([TextCellValue('Report Generated'),TextCellValue(DateFormat('dd-MM-yyyy HH:mm').format(DateTime.now()))]);
   dash.appendRow([TextCellValue('Total Trips'),IntCellValue(trips.length)]);
   dash.appendRow([TextCellValue('Total KM'),DoubleCellValue(totalKm)]);
   dash.appendRow([TextCellValue('TOTAL HRS'),TextCellValue(sumHours(trips.map((t)=>t.totalHrs)))]);
   dash.appendRow([TextCellValue('EX. HRS'),TextCellValue(sumHours(trips.map((t)=>t.exHrs)))]);
   dash.appendRow([TextCellValue('Extra Hours Rule'),TextCellValue('Extra hours calculated after 16 duty hours.')]);dash.setColumnWidth(0,28);dash.setColumnWidth(1,42);
   final sheet=book['Vehicle Log'];const h=['SR NO','DATE','VEHICLE NUMBER','FROM TIME','TO TIME','FROM KM','TO KM','TOTAL KM','BRIEF DETAILS OF JOURNEY','NAME OF USER','TOTAL HRS','EX. HRS','DRIVER'];sheet.appendRow(h.map((x)=>TextCellValue(x)).toList());
   for(var i=0;i<trips.length;i++){final t=trips[i];sheet.appendRow([IntCellValue(i+1),TextCellValue(t.date),TextCellValue(t.vehicle),TextCellValue(t.fromTime),TextCellValue(t.toTime),DoubleCellValue(t.fromKm),DoubleCellValue(t.toKm),DoubleCellValue(t.km),TextCellValue(t.journey),TextCellValue(t.user),TextCellValue(t.totalHrs),TextCellValue(t.exHrs),TextCellValue(t.driver)]);}
   for(var i=0;i<h.length;i++){sheet.setColumnWidth(i,[9,14,20,12,12,12,12,12,38,22,12,12,20][i].toDouble());}
   final bytes=book.encode();if(bytes==null)throw Exception('Excel create failed');final dir=await getTemporaryDirectory();final name='ABHILASH_TRANSPORT_Logbook_${DateFormat('yyyyMMdd_HHmm').format(DateTime.now())}.xlsx';final file=File(p.join(dir.path,name));await file.writeAsBytes(bytes,flush:true);
   await Share.shareXFiles([XFile(file.path,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',name:name)],subject:'ABHILASH TRANSPORT Logbook',text:'Vehicle logbook Excel report');
  }catch(e){note('Excel export failed: $e');}finally{if(mounted)setState(()=>busy=false);}
 }
 Widget tile(String label,String value,IconData icon,Color color)=>Card(color:color,child:Padding(padding:const EdgeInsets.all(13),child:Column(crossAxisAlignment:CrossAxisAlignment.start,children:[Icon(icon,color:navy),const SizedBox(height:6),Text(value,style:const TextStyle(fontSize:22,fontWeight:FontWeight.bold,color:navy)),Text(label)])));
 @override Widget build(BuildContext c)=>Scaffold(appBar:AppBar(backgroundColor:Colors.white,title:const BrandLogo(height:48),centerTitle:true,actions:[IconButton(onPressed:refresh,icon:const Icon(Icons.refresh))]),floatingActionButton:FloatingActionButton.extended(onPressed:()async{final ok=await Navigator.push<bool>(c,MaterialPageRoute(builder:(_)=>const Entry()));if(ok==true)refresh();},icon:const Icon(Icons.add),label:const Text('New Entry')),body:RefreshIndicator(onRefresh:refresh,child:ListView(padding:const EdgeInsets.all(12),children:[
  Container(padding:const EdgeInsets.all(18),decoration:BoxDecoration(gradient:const LinearGradient(colors:[navy,blue]),borderRadius:BorderRadius.circular(18)),child:Column(crossAxisAlignment:CrossAxisAlignment.start,children:[const Text('WELCOME TO',style:TextStyle(color:Colors.white70)),const Text('ABHILASH TRANSPORT',style:TextStyle(color:Colors.white,fontSize:22,fontWeight:FontWeight.w900)),Text('Vehicle Logbook • ${DateFormat('dd MMM yyyy').format(DateTime.now())}',style:const TextStyle(color:Colors.white70))])),
  GridView.count(shrinkWrap:true,physics:const NeverScrollableScrollPhysics(),crossAxisCount:2,childAspectRatio:1.6,children:[tile('Total Trips','${trips.length}',Icons.directions_car,const Color(0xffe3efff)),tile('Total KM',totalKm.toStringAsFixed(1),Icons.route,const Color(0xffdff7e9)),tile('TOTAL HRS',sumHours(trips.map((t)=>t.totalHrs)),Icons.schedule,const Color(0xfffff0d4)),tile('EX. HRS',sumHours(trips.map((t)=>t.exHrs)),Icons.more_time,const Color(0xffffe4e4))]),
  const Text('Recent Trips',style:TextStyle(fontSize:19,fontWeight:FontWeight.bold,color:navy)),
  if(trips.isEmpty)const Card(child:Padding(padding:EdgeInsets.all(24),child:Text('No trips yet. Tap New Entry to begin.')))
  else ...trips.take(10).map((t)=>Card(child:ListTile(leading:const Icon(Icons.directions_car,color:navy),title:Text('${t.vehicle} • ${t.date}'),subtitle:Text('${t.fromTime} → ${t.toTime}\\nTOTAL HRS ${t.totalHrs} | EX. HRS ${t.exHrs}\\n${t.journey}'),isThreeLine:true,trailing:Text('${t.km.toStringAsFixed(0)} KM')))),
  FilledButton.icon(onPressed:busy?null:shareExcel,icon:const Icon(Icons.share),label:Text(busy?'Preparing Excel…':'Export & Share Excel')),
  const Center(child:Text('Excel includes Dashboard and Vehicle Log sheets',style:TextStyle(fontSize:12)))
 ])));
}

class Entry extends StatefulWidget {const Entry({super.key});@override State<Entry> createState()=>_EntryState();}
class _EntryState extends State<Entry>{
 final form=GlobalKey<FormState>();late final date=TextEditingController(text:DateFormat('yyyy-MM-dd').format(DateTime.now()));
 final vehicle=TextEditingController(text:'MH12ZU6903'),fromTime=TextEditingController(),toTime=TextEditingController(),fromKm=TextEditingController(),toKm=TextEditingController(),journey=TextEditingController(),user=TextEditingController(),driver=TextEditingController();bool saving=false;
 String? req(String? v)=>v==null||v.trim().isEmpty?'Required':null;
 int? parseClock(String s){final m=RegExp(r'^(\d{1,2}):(\d{2})$').firstMatch(s.trim());if(m==null)return null;final h=int.tryParse(m.group(1)!),n=int.tryParse(m.group(2)!);if(h==null||n==null||h>23||n>59)return null;return h*60+n;}
 int? get worked{final a=parseClock(fromTime.text),b=parseClock(toTime.text);if(a==null||b==null)return null;var n=b-a;if(n<0)n+=1440;return n;}
 String fmt(int n)=>'${n~/60}:${(n%60).toString().padLeft(2,'0')}';
 String get totalHrs=>worked==null?'—':fmt(worked!);String get exHrs=>worked==null?'—':fmt((worked!-960).clamp(0,1440));
 double? get km{final a=double.tryParse(fromKm.text),b=double.tryParse(toKm.text);return a!=null&&b!=null&&b>=a?b-a:null;}
 Widget field(String label,TextEditingController ctrl,{TextInputType? type})=>TextFormField(controller:ctrl,keyboardType:type,decoration:InputDecoration(labelText:label,border:const OutlineInputBorder()),validator:req,onChanged:(_)=>setState((){}));
 Future<void> chooseDate()async{final d=await showDatePicker(context:context,initialDate:DateTime.tryParse(date.text)??DateTime.now(),firstDate:DateTime(2020),lastDate:DateTime(2100));if(d!=null)setState(()=>date.text=DateFormat('yyyy-MM-dd').format(d));}
 Future<void> chooseTime(TextEditingController ctrl)async{final t=await showTimePicker(context:context,initialTime:TimeOfDay.now());if(t!=null)setState(()=>ctrl.text='${t.hour.toString().padLeft(2,'0')}:${t.minute.toString().padLeft(2,'0')}');}
 Future<void> save()async{if(!form.currentState!.validate())return;final a=double.tryParse(fromKm.text),b=double.tryParse(toKm.text);if(a==null||b==null||b<a){ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content:Text('Ending KM cannot be less than Starting KM')));return;}if(worked==null){ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content:Text('Please select valid From and To Time')));return;}setState(()=>saving=true);try{await DB.instance.add(Trip(date:date.text,vehicle:vehicle.text.trim().toUpperCase(),fromTime:fromTime.text,toTime:toTime.text,fromKm:a,toKm:b,journey:journey.text.trim(),user:user.text.trim(),driver:driver.text.trim(),totalHrs:totalHrs,exHrs:exHrs));if(mounted)Navigator.pop(context,true);}catch(e){if(mounted)ScaffoldMessenger.of(context).showSnackBar(SnackBar(content:Text('Save failed: $e')));}finally{if(mounted)setState(()=>saving=false);}}
 @override void dispose(){for(final x in [date,vehicle,fromTime,toTime,fromKm,toKm,journey,user,driver]){x.dispose();}super.dispose();}
 @override Widget build(BuildContext c)=>Scaffold(appBar:AppBar(backgroundColor:navy,foregroundColor:Colors.white,title:const Text('Add Trip Entry')),body:Form(key:form,child:ListView(padding:const EdgeInsets.all(16),children:[
  const BrandLogo(height:64),InkWell(onTap:chooseDate,child:InputDecorator(decoration:const InputDecoration(labelText:'Date *',border:OutlineInputBorder(),suffixIcon:Icon(Icons.calendar_month)),child:Text(date.text))),const SizedBox(height:12),
  field('Vehicle Number *',vehicle),const SizedBox(height:12),
  Row(children:[Expanded(child:InkWell(onTap:()=>chooseTime(fromTime),child:InputDecorator(decoration:const InputDecoration(labelText:'From Time *',border:OutlineInputBorder()),child:Text(fromTime.text.isEmpty?'Select time':fromTime.text)))),const SizedBox(width:10),Expanded(child:InkWell(onTap:()=>chooseTime(toTime),child:InputDecorator(decoration:const InputDecoration(labelText:'To Time *',border:OutlineInputBorder()),child:Text(toTime.text.isEmpty?'Select time':toTime.text))))]),const SizedBox(height:12),
  Row(children:[Expanded(child:field('Starting KM *',fromKm,type:const TextInputType.numberWithOptions(decimal:true))),const SizedBox(width:10),Expanded(child:field('Ending KM *',toKm,type:const TextInputType.numberWithOptions(decimal:true)))]),const SizedBox(height:8),
  Text('TOTAL KM: ${km?.toStringAsFixed(1)??'—'}',style:const TextStyle(fontWeight:FontWeight.bold,color:navy)),const SizedBox(height:12),
  field('Brief Details of Journey *',journey),const SizedBox(height:12),field('Name of User *',user),const SizedBox(height:12),field('Driver Name *',driver),const SizedBox(height:14),
  Row(children: [Expanded(child: InputDecorator(decoration: const InputDecoration(labelText: 'TOTAL HRS'), child: Text(totalHrs, style: const TextStyle(fontSize: 22, fontWeight: FontWeight.bold)))), const SizedBox(width: 10), Expanded(child: InputDecorator(decoration: const InputDecoration(labelText: 'EX. HRS (AFTER 16 HRS)'), child: Text(exHrs, style: const TextStyle(fontSize: 22, fontWeight: FontWeight.bold, color: Color(0xffb42318))))) ]), const SizedBox(height: 20),
  FilledButton.icon(onPressed:saving?null:save,icon:const Icon(Icons.save),label:Text(saving?'Saving…':'Save Entry'))
 ])));
}
