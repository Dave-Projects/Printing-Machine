// ignore_for_file: use_build_context_synchronously

import 'package:cap102/pages/homepage.dart';
import 'package:cap102/pages/sales.dart';
import 'package:cap102/pages/signIn.dart';
import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:flutter/material.dart';

class Account extends StatefulWidget {
  const Account({super.key});

  @override
  State<Account> createState() => _AccountState();
}

class _AccountState extends State<Account> {
  int myIndex = 0;

  // Reference to Firestore collection
  final CollectionReference database =
      FirebaseFirestore.instance.collection("users");

  // Controllers for text fields
  final TextEditingController nameController = TextEditingController();
  final TextEditingController contactController = TextEditingController();
  final TextEditingController emailController = TextEditingController();

  bool isEditing = false; // Track edit mode

  @override
  void initState() {
    super.initState();
    fetchData();
  }

  Future<void> fetchData() async {
    try {
      print('Fetching document from Firestore...');
      DocumentSnapshot snapshot =
          await database.doc("BVmRKnW9ylJZ6bZFx5it").get();

      if (snapshot.exists) {
        print('Document exists: ${snapshot.data()}');
        setState(() {
          nameController.text = snapshot['Full Name']?.toString() ?? '';
          contactController.text = snapshot['Contact']?.toString() ?? '';
          emailController.text = snapshot['Email']?.toString() ?? '';
        });
      } else {
        print('Document does not exist');
      }
    } catch (e) {
      print('Error fetching data: $e');
    }
  }

  Future<void> saveData() async {
    try {
      await database.doc("BVmRKnW9ylJZ6bZFx5it").update({
        'Full Name': nameController.text,
        'Contact': contactController.text,
        'Email': emailController.text,
      });
      print("Data saved successfully");
    } catch (e) {
      print("Error saving data: $e");
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
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
          if (index == 2) {
            Navigator.push(context,
                MaterialPageRoute(builder: (context) => Sales()));
          }
        },
        currentIndex: 1,
        items: const [
          BottomNavigationBarItem(icon: Icon(Icons.home), label: 'Dashboard'),
          BottomNavigationBarItem(
              icon: Icon(Icons.account_circle), label: 'Account'),
          BottomNavigationBarItem(
              icon: Icon(Icons.shop_rounded), label: 'Sales'),
        ],
      ),
      appBar: AppBar(
        title: const Text("Account"),
        centerTitle: true,
        backgroundColor: const Color(0xffaf0b00),
        automaticallyImplyLeading: false,
        foregroundColor: Colors.white,
        actions: [
          Padding(
            padding: const EdgeInsets.symmetric(horizontal: 30),
            child: GestureDetector(
              onTap: () {
                setState(() {
                  isEditing = !isEditing; // Toggle editing mode
                  if (!isEditing) {
                    saveData(); // Save data when toggling back to read-only
                  }
                });
              },
              child: Text(
                isEditing ? "Save" : "Edit",
                style: TextStyle(fontSize: 18, color: Colors.white),
              ),
            ),
          ),
        ],
      ),
      body: Padding(
        padding: const EdgeInsets.symmetric(vertical: 30),
        child: Center(
          child: Column(
            children: [
              const Icon(
                Icons.account_circle_rounded,
                color: Colors.deepPurple,
                size: 100,
              ),
              const SizedBox(height: 20),
              const Text(
                "Account Details",
                style: TextStyle(fontSize: 20),
              ),
              const SizedBox(height: 10),
              textField('Name: ', nameController, !isEditing),
              textField('Contact No: ', contactController, !isEditing),
              textField('Email: ', emailController, !isEditing),
              const SizedBox(height: 10),
              ElevatedButton(
                onPressed: () {
                  signOut(context);
                },
                style: ElevatedButton.styleFrom(shadowColor: null),
                child: Text("Logout"),
              )
            ],
          ),
        ),
      ),
    );
  }

  Widget textField(
      String label, TextEditingController controller, bool readonly) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 15, vertical: 5),
      child: TextField(
        controller: controller,
        decoration: InputDecoration(
          labelText: label,
          labelStyle: const TextStyle(fontSize: 20, color: Colors.blue),
          border: OutlineInputBorder(
            borderRadius: BorderRadius.circular(20),
          ),
        ),
        style: const TextStyle(fontSize: 20),
        readOnly: readonly,
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
}
