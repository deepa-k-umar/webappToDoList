# Monthly Todo for Android

This is a native Android version of the todo tracker. It stores tasks locally on the phone and works without Python, Excel, or an internet connection.

## Features

- Add and delete tasks
- View tasks by month
- Mark tasks complete
- Automatically create the next monthly task when completed
- Task data persists between app launches

## Build and install

1. Install Android Studio with the Android SDK and SDK Platform 35.
2. Open this `android-app` folder in Android Studio and let Gradle sync.
3. Enable Developer options and USB debugging on the phone, then connect it by USB.
4. Run the `app` configuration from Android Studio, or choose **Build > Build APK(s)**.

The debug APK will be created at:

`app/build/outputs/apk/debug/app-debug.apk`

You can install it from a terminal after enabling USB debugging:

```text
adb install -r app/build/outputs/apk/debug/app-debug.apk
```
