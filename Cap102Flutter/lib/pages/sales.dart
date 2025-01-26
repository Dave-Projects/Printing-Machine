import 'dart:async';
import 'package:cap102/pages/account.dart';
import 'package:cap102/pages/homepage.dart';
import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';

class Sales extends StatefulWidget {
  const Sales({super.key});

  @override
  State<Sales> createState() => _SalesState();
}

class _SalesState extends State<Sales> {
  int myIndex = 0;
  String formattedDate = '';
  Timer? timer;

  final CollectionReference fetchData = FirebaseFirestore.instance.collection("sales");

  @override
  void initState() {
    super.initState();
    updateDate();
    timer = Timer.periodic(const Duration(seconds: 60), (Timer t) => updateDate());
  }

  void updateDate() {
    final now = DateTime.now();
    setState(() {
      formattedDate = DateFormat('yyyy-MM-dd').format(now); // Update the date only
    });
  }

  @override
  void dispose() {
    timer?.cancel(); // Cancel timer when widget is disposed
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: const Color(0xffaf0b00),
      appBar: AppBar(
        title: const Text(
          "Sales",
          style: TextStyle(fontSize: 25),
        ),
        centerTitle: true,
        backgroundColor: const Color.fromARGB(255, 133, 23, 15),
        foregroundColor: Colors.white,
        titleTextStyle: const TextStyle(fontSize: 18),
      ),
      bottomNavigationBar: BottomNavigationBar(
        onTap: (index) {
          setState(() {
            myIndex = index;
          });
          if (index == 0) {
            Navigator.push(
              context,
              MaterialPageRoute(builder: (context) => Homepage()),
            );
          }
          if (index == 1) {
            Navigator.push(
              context,
              MaterialPageRoute(builder: (context) => Account()),
            );
          }
        },
        currentIndex: 2,
        items: const [
          BottomNavigationBarItem(icon: Icon(Icons.home), label: 'Dashboard'),
          BottomNavigationBarItem(
              icon: Icon(Icons.account_circle), label: 'Account'),
          BottomNavigationBarItem(
              icon: Icon(Icons.shop_rounded), label: 'Sales'),
        ],
      ),
      body: Column(
        children: [
          // Date and Total Sales
          Container(
            height: 250,
            width: double.infinity,
            color: const Color(0xffaf0b00),
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                Text(
                  formattedDate,
                  style: const TextStyle(fontSize: 25, color: Colors.white),
                ),
                const SizedBox(height: 30),
                _buildTotalSales(),
              ],
            ),
          ),
          // Sales History
          Expanded(
            child: Container(
              decoration: const BoxDecoration(
                color: Colors.white,
                borderRadius: BorderRadius.only(
                  topLeft: Radius.circular(30),
                  topRight: Radius.circular(30),
                ),
              ),
              child: Column(
                children: [
                  const SizedBox(height: 15),
                  const Text(
                    "Sales History",
                    style: TextStyle(fontSize: 25),
                  ),
                  Expanded(
                    child: _buildSalesHistory(), // Make sure this is scrollable
                  ),
                ],
              ),
            ),
          ),
        ],
      ),
    );
  }

  /// Widget to build the Total Sales using a StreamBuilder
  Widget _buildTotalSales() {
    return StreamBuilder<QuerySnapshot>(
      stream: fetchData.snapshots(),
      builder: (context, snapshot) {
        if (snapshot.connectionState == ConnectionState.waiting) {
          return const CircularProgressIndicator(); // Show loading indicator
        }
        if (snapshot.hasError) {
          return Text(
            "Error: ${snapshot.error}",
            style: const TextStyle(color: Colors.white),
          );
        }
        if (snapshot.hasData && snapshot.data!.docs.isNotEmpty) {
          double totalSales = 0.0;
          for (var doc in snapshot.data!.docs) {
            final amount = double.tryParse(doc['amount'].toString()) ?? 0.0;
            totalSales += amount;
          }
          return Text(
            "Total Sales: ₱${totalSales.toStringAsFixed(2)}",
            style: const TextStyle(
                fontSize: 30, color: Colors.white, fontWeight: FontWeight.w600),
          );
        } else {
          return const Text(
            "Total Sales: ₱0.00",
            style: TextStyle(
                fontSize: 30, color: Colors.white,),
          );
        }
      },
    );
  }

  /// Widget to build the Sales History using a StreamBuilder
  Widget _buildSalesHistory() {
    return StreamBuilder<QuerySnapshot>(
      stream: fetchData.snapshots(),
      builder: (context, snapshot) {
        if (snapshot.connectionState == ConnectionState.waiting) {
          return const Center(
            child: CircularProgressIndicator(),
          );
        }
        if (snapshot.hasError) {
          return Center(
            child: Text("Error fetching data: ${snapshot.error}"),
          );
        }
        if (snapshot.hasData && snapshot.data!.docs.isNotEmpty) {
          return ListView.builder(
            // Ensure ListView.builder is scrollable
            shrinkWrap: true,
            itemCount: snapshot.data!.docs.length,
            itemBuilder: (context, index) {
              final DocumentSnapshot doc = snapshot.data!.docs[index];
              return ListTile(
                title: Text(
                  "Amount: ₱" + doc['amount'].toString(),
                  style: const TextStyle(
                    fontWeight: FontWeight.bold,
                    fontSize: 20,
                    color: Colors.black,
                  ),
                ),
                subtitle: Text(
                  "Date: " + doc['date'].toString(),
                  style: const TextStyle(color: Colors.black),
                ),
              );
            },
          );
        } else {
          return const Center(
            child: Text("No sales data available."),
          );
        }
      },
    );
  }
}
