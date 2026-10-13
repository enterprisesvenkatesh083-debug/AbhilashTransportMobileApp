ABHILASH TRANSPORT Mobile Logbook

Features:
- Android mobile entry form: date, vehicle, times, starting/ending KM, journey, user, driver, total/extra hours.
- Total KM automatically calculates.
- Entries are stored on the phone.
- Export & Share Excel creates an .xlsx workbook and opens Android's share menu (WhatsApp, Gmail, Drive, etc.).

Important: This is Flutter source code, not an already-built APK. Entries stay on this phone until you share/export the Excel file. It does not automatically sync to office Excel.

Build APK:
1. Install Flutter SDK and Android Studio on a computer.
2. Create a Flutter project scaffold using: flutter create --platforms android .
   (Run this command in a separate empty folder, then copy this package's pubspec.yaml and lib/main.dart into it.)
3. Run: flutter pub get
4. Run: flutter build apk --release
5. APK path: build/app/outputs/flutter-apk/app-release.apk

Then send the APK to the driver's Android phone and install it. Android may ask permission to install unknown apps.
