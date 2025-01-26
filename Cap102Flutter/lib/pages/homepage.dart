// ignore_for_file: prefer_const_constructors, prefer_const_literals_to_create_immutables, sort_child_properties_last, use_build_context_synchronously

import 'dart:async';

import 'package:cap102/pages/account.dart';
import 'package:cap102/pages/sales.dart';
import 'package:cap102/pages/signIn.dart';
import 'package:cap102/services/notification.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:firebase_database/firebase_database.dart';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'package:percent_indicator/circular_percent_indicator.dart';

class Homepage extends StatefulWidget {
  const Homepage({super.key});

  @override
  State<Homepage> createState() => _HomepageState();
}

class _HomepageState extends State<Homepage> {
  final user = FirebaseAuth.instance.currentUser!;
  String realtimeValue = '0';
  double percentValue = 0.0;
  String realtimeValue1 = '0';
  double percentValue1 = 0.0;
  String formattedDate = '';
  String formattedTime = '';
  Timer? timer;

  double totalSales = 0.0; // Variable to store total sales
  bool notificationSent = false; // Track if notification has been sent

  // ignore: prefer_typing_uninitialized_variables
  var height, width;

  int myIndex = 0;
  bool showPrinter1 = true; // New variable to toggle between printers

  @override
  void initState() {
    super.initState();

    // Initialize date-time and timer to update every second
    updateDateTime();
    timer = Timer.periodic(Duration(seconds: 1), (Timer t) => updateDateTime());

    // Initialize Firebase Realtime Database listeners
    DatabaseReference ref = FirebaseDatabase.instance.ref().child('weight1');
    ref.onValue.listen(
      (event) {
        setState(() {
          realtimeValue = event.snapshot.value.toString();
          percentValue = calculatePercent(realtimeValue);
          checkAndNotify(realtimeValue, "Printer 1's Paper Tray Weight");
        });
      },
    );
    DatabaseReference ref1 = FirebaseDatabase.instance.ref().child('weight2');
    ref1.onValue.listen(
      (event) {
        setState(() {
          realtimeValue1 = event.snapshot.value.toString();
          percentValue1 = calculatePercent(realtimeValue1);
          checkAndNotify(realtimeValue1, "Printer 2's Paper Tray Weight");
        });
      },
    );

    // Initialize the total sales calculation
    fetchTotalSales();
  }

  /// Function to fetch total sales from Firestore
  void fetchTotalSales() {
    FirebaseFirestore.instance
        .collection('sales')
        .snapshots()
        .listen((snapshot) {
      double total = 0.0;
      for (var doc in snapshot.docs) {
        final amount = double.tryParse(doc['amount'].toString()) ?? 0.0;
        total += amount;
      }
      setState(() {
        totalSales = total;
      });
    });
  }

  void updateDateTime() {
    final now = DateTime.now();
    setState(() {
      formattedDate = DateFormat('yyyy-MM-dd').format(now); // Format only date
      formattedTime = DateFormat('HH:mm:ss').format(now); // Format only time
    });
  }

  void checkAndNotify(String value, String printerName) {
    double weightValue = double.tryParse(value) ?? 0.0;
    if (weightValue <= 20) {
      // Trigger a notification if the weight is 20 or below
      LocalNotifications.showSimpleNotification(
        title: 'Low Paper Tray Weight',
        body: '$printerName is running low on paper (Weight: $value grams).',
        payload: 'Check $printerName',
      );
    }
  }

  @override
  void dispose() {
    timer?.cancel(); // Cancel the timer when the widget is disposed
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    height = MediaQuery.of(context).size.height;
    width = MediaQuery.of(context).size.width;
    return Scaffold(
        backgroundColor: Color(0xffFCCCCC),
        bottomNavigationBar: BottomNavigationBar(
            onTap: (index) {
              setState(() {
                myIndex = index;
              });
              if (index == 1) {
                Navigator.push(
                  context,
                  MaterialPageRoute(builder: (context) => Account()),
                );
              }
              if (index == 2) {
                Navigator.push(
                    context, MaterialPageRoute(builder: (context) => Sales()));
              }
            },
            currentIndex: 0,
            items: const [
              BottomNavigationBarItem(
                  icon: Icon(Icons.home), label: 'Dashboard'),
              BottomNavigationBarItem(
                  icon: Icon(Icons.account_circle), label: 'Account'),
              BottomNavigationBarItem(
                  icon: Icon(Icons.shop_rounded), label: 'Sales'),
            ]),
        body: Container(
          decoration: BoxDecoration(
            color: Color(0xffaf0b00),
          ),
          height: 800,
          width: 400,
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Padding(
                padding:
                    const EdgeInsets.symmetric(vertical: 55, horizontal: 15),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        RichText(
                            text: TextSpan(children: <TextSpan>[
                          TextSpan(
                              text: "Hello, ",
                              style: TextStyle(
                                  color: Colors.white,
                                  fontWeight: FontWeight.w700,
                                  fontSize: 25)),
                          TextSpan(
                              text: "Administrator!",
                              style: TextStyle(
                                  color: Colors.white,
                                  fontWeight: FontWeight.w300,
                                  fontSize: 25)),
                        ])),
                      ],
                    ),
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          formattedDate,
                          style: TextStyle(fontSize: 20, color: Colors.white),
                        ),
                        Text(
                          formattedTime,
                          style: TextStyle(fontSize: 18, color: Colors.white),
                        ),
                        SizedBox(height: 15),
                        Text(
                          "Total Sales",
                          style: TextStyle(fontSize: 30, color: Colors.white),
                        ),
                        Row(
                          children: [
                            Text(
                              "₱ ${totalSales.toStringAsFixed(2)}",
                              style:
                                  TextStyle(fontSize: 40, color: Colors.white),
                            ),
                            Spacer(),
                          ],
                        )
                      ],
                    )
                  ],
                ),
              ),
              Container(
                height: 459.9,
                width: 400,
                decoration: BoxDecoration(
                  borderRadius: BorderRadius.only(
                      topLeft: Radius.circular(30),
                      topRight: Radius.circular(30)),
                  color: Colors.white,
                ),
                child: Column(
                  children: [
                    SizedBox(height: 15),
                    Text(
                      "Printer Status",
                      style: TextStyle(fontSize: 20),
                    ),
                    SizedBox(height: 15),
                    // Use `showPrinter1` to toggle between printers
                    showPrinter1
                        ? printer(percentValue, realtimeValue,
                            "Printer 1's Paper Tray Weight")
                        : printer(percentValue1, realtimeValue1,
                            "Printer 2's Paper Tray Weight"),
                    SizedBox(height: 15),
                    ElevatedButton(
                      onPressed: () => setState(() {
                        showPrinter1 = !showPrinter1; // Toggle between printers
                      }),
                      child: Text(
                          showPrinter1 ? "View Printer 2" : "View Printer 1"),
                      style: ElevatedButton.styleFrom(
                          foregroundColor: Colors.deepPurple,
                          shape: RoundedRectangleBorder(
                              borderRadius: BorderRadius.circular(20)),
                          shadowColor: null),
                    )
                  ],
                ),
              )
            ],
          ),
        ));
  }
}

Widget printer(double percentValue, String realtimeValue, String title) {
  double weightValue = double.tryParse(realtimeValue) ?? 0.0;
  Color progressColor = weightValue <= 20 ? Colors.red : Colors.deepPurple;

  return CircularPercentIndicator(
    radius: 120,
    lineWidth: 15,
    percent: percentValue,
    animation: true,
    animateFromLastPercent: true,
    animationDuration: 500,
    progressColor: progressColor,
    backgroundColor: Colors.deepPurple.shade100,
    circularStrokeCap: CircularStrokeCap.round,
    center: Text(
      "$realtimeValue Grams",
      style: TextStyle(fontSize: 20),
    ),
    footer: Padding(
      padding: const EdgeInsets.symmetric(vertical: 10),
      child: Text(
        title, // Display the printer-specific title
        style: TextStyle(fontSize: 20, fontWeight: FontWeight.w400),
      ),
    ),
  );
}

Future<void> signOut(BuildContext context) async {
  await FirebaseAuth.instance.signOut();
  showDialog(
    context: context,
    builder: (context) {
      return const Center(
        child: CircularProgressIndicator(),
      );
    },
  );
  await Future.delayed(const Duration(seconds: 1));
  Navigator.pushReplacement(
    context,
    MaterialPageRoute(
      builder: (BuildContext context) => SignIn(),
    ),
  );
}

double calculatePercent(String value) {
  double? numericValue = double.tryParse(value);
  if (numericValue != null) {
    return numericValue / 100; // Convert to a percent between 0 and 1
  }
  return 0.0; // Default to 0 if parsing fails
}
