// ignore_for_file: prefer_const_constructors, file_names, use_build_context_synchronously

import 'package:cap102/pages/signIn.dart';
import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:flutter/material.dart';

class SignUp extends StatefulWidget {
  const SignUp({super.key});

  @override
  State<SignUp> createState() => _SignUpState();
}

class _SignUpState extends State<SignUp> {
  final TextEditingController _emailController = TextEditingController();
  final TextEditingController _passwordController = TextEditingController();
  final TextEditingController _confirmPasswordController =
      TextEditingController();
  final TextEditingController _nameController = TextEditingController();
  final TextEditingController _contactController = TextEditingController();

  final double height = 10;

  final double radius = 20;

  //static List<Map<String, String>> userAccounts = [];
  @override
  Widget build(BuildContext context) {
    return SafeArea(
      child: Scaffold(
       backgroundColor: Color.fromARGB(255, 133, 23, 15),
        body: Center(
          child: Column(
            children: [
              SizedBox(height: 30),
              Text(
                "PrintVendo",
                style: TextStyle(
                    fontSize: 60,
                    color: Colors.white,
                    fontWeight: FontWeight.bold),
              ),
              Padding(
                padding: const EdgeInsets.symmetric(vertical: 1.0),
                child: Text(
                  "Sign Up",
                  style: TextStyle(
                      fontSize: 25,
                      color: Colors.white,
                      fontWeight: FontWeight.bold),
                ),
              ),
              SizedBox(height: 20),
              Container(
                height: 530,
                width: 335,
                decoration: BoxDecoration(
                    borderRadius: BorderRadius.circular(radius),
                    color: Colors.white),
                child: Padding(
                  padding:
                      const EdgeInsets.symmetric(horizontal: 30, vertical: 25),
                  child: Column(
                    children: [
                      // Username textfield
                      Text(
                        "Create an account to supervise your system.",
                        style: TextStyle(fontSize: 15, color: Colors.black),
                        textAlign: TextAlign.center,
                      ),
                      SizedBox(height: 10),
                      textField(_nameController, 'Full Name', false),
                      SizedBox(height: 10),
                      textField(_contactController, 'Contact No.', false),
                      SizedBox(height: 10),
                      textField(_emailController, 'Email', false),
                      SizedBox(height: 10),
                      // Password textfield
                      textField(_passwordController, 'Password', true),
                      SizedBox(height: 10),
                      // Confirm Password textfield
                      textField(
                          _confirmPasswordController, 'Confirm Password', true),
                      SizedBox(
                        height: radius,
                      ),
                      ElevatedButton(
                        onPressed: () {
                          signUp(context);
                        },
                        style: ElevatedButton.styleFrom(
                            backgroundColor: Color(0xffDAB4DD),
                            shape: RoundedRectangleBorder(
                                borderRadius: BorderRadius.circular(radius))),
                        child: Text(
                          "Create Account",
                          style: TextStyle(color: Colors.white, fontSize: 15),
                        ),
                      ),
                      //Whitespace
                      SizedBox(
                        height: 15,
                      ),
                      Row(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          Padding(
                            padding:
                                const EdgeInsets.symmetric(horizontal: 10.0),
                            child: Text(
                              "Already have an account?",
                              style: TextStyle(
                                color: Colors.black,
                              ),
                            ),
                          ),
                          GestureDetector(
                            onTap: () {
                              Navigator.push(
                                context,
                                MaterialPageRoute(
                                    builder: (context) => SignIn()),
                              );
                            },
                            child: Text(
                              "Sign In",
                              style: TextStyle(
                                color: Color(0xff7E3185),
                                fontSize: 15,
                              ),
                            ),
                          )
                        ],
                      ),
                    ],
                  ),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  void signUp(BuildContext context) async {
    final email = _emailController.text;
    final password = _passwordController.text;
    final confirmPassword = _confirmPasswordController.text;

/*
    showDialog(
      context: context,
      builder: (context) {
        return const Center(
          child: CircularProgressIndicator(),
        );
      },
    );
*/
    try {
      if (password == confirmPassword) {
        await FirebaseAuth.instance
            .createUserWithEmailAndPassword(email: email, password: password);
        addUserDetails(
            _nameController.text.trim(),
            int.parse(_contactController.text.trim()),
            _emailController.text.trim());

        showMessage(context, "Account successfully created!");
        clear();
      } else {
        showMessage(context, "Passwords don't match!");
      }
    } on FirebaseAuthException catch (e) {
      if (e.code == 'email-already-in-use') {
        showMessage(context, "Email already exists!");
      } else if (e.code == 'weak-password') {
        showMessage(context, "Password length must be at least 6 characters!");
      }
    }
  }

  Future addUserDetails(String fullName, int contactNo, String email) async {
    await FirebaseFirestore.instance
        .collection('users')
        .add({'Full Name': fullName, 'Contact': contactNo, 'Email': email});
  }

  void clear() {
    _nameController.clear();
    _contactController.clear();
    _emailController.clear();
    _passwordController.clear();
    _confirmPasswordController.clear();
  }

  void showMessage(BuildContext context, String message) {
    showDialog(
      context: context,
      builder: (context) {
        return AlertDialog(
          shape:
              RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
          title: Center(
            child: Text(
              message,
              style: const TextStyle(color: Colors.black, fontSize: 25),
            ),
          ),
          actions: [
            MaterialButton(
              onPressed: () {
                Navigator.pop(context);
              },
              child: const Text("Okay"),
            )
          ],
        );
      },
    );
  }

  Widget textField(
      TextEditingController controller, String label, bool hideText) {
    return TextFormField(
      controller: controller,
      decoration: InputDecoration(
        enabledBorder: OutlineInputBorder(
          borderSide: BorderSide(color: Colors.black),
          borderRadius: BorderRadius.circular(20),
        ),
        focusedBorder: OutlineInputBorder(
          borderSide: BorderSide(color: Colors.grey.shade400),
          borderRadius: BorderRadius.circular(20),
        ),
        fillColor: Colors.white,
        labelText: label,
        filled: true,
      ),
      obscureText: hideText,
    );
  }
}
