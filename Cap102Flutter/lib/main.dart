// ignore_for_file: prefer_const_constructors, use_key_in_widget_constructors

import 'package:cap102/pages/account.dart';
import 'package:cap102/pages/homepage.dart';
import 'package:cap102/pages/index.dart';
import 'package:cap102/pages/sales.dart';
import 'package:cap102/pages/signIn.dart';
import 'package:cap102/pages/signUp.dart';
import 'package:cap102/services/notification.dart';
import 'package:firebase_core/firebase_core.dart';
import 'package:flutter/material.dart';

import 'firebase_options.dart';

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await LocalNotifications.init();
  await Firebase.initializeApp(
    options: DefaultFirebaseOptions.currentPlatform,
  );
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: SignIn(), // Initial route
      debugShowCheckedModeBanner: false,
      routes: {
        '/signIn': (context) => SignIn(),
        '/signUp': (context) => SignUp(),
        '/homepage': (context) => Homepage(),
        '/index': (context) => IntroPage(),
        '/account': (context) => Account(),
        'sales': (context) => Sales(),
      },
    );
  }
}
