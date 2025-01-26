// ignore_for_file: prefer_const_constructors, sort_child_properties_last, file_names, use_build_context_synchronously

import 'package:cap102/pages/signUp.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:flutter/material.dart';

class SignIn extends StatefulWidget {
  const SignIn({super.key});

  @override
  State<SignIn> createState() => _SignInState();
}

class _SignInState extends State<SignIn> {
  final TextEditingController _emailController = TextEditingController();
  final TextEditingController _passwordController = TextEditingController();

  final double height = 20;

  @override
  Widget build(BuildContext context) {
    return SafeArea(
        child: Scaffold(
      backgroundColor: Color.fromARGB(255, 133, 23, 15),
      body: Center(
        child: Column(
          children: [
            //Whitespace
            SizedBox(height: 100),

            Text("PrintVendo",
                style: TextStyle(
                    color: Colors.white,
                    fontSize: 60,
                    fontWeight: FontWeight.bold)),
            Text("Sign In",
                style: TextStyle(
                    color: Colors.white,
                    fontSize: 20,
                    fontWeight: FontWeight.bold)),
            //Whitespace
            SizedBox(height: 60),

            Container(
              height: 320,
              width: 335,
              decoration: BoxDecoration(
                  borderRadius: BorderRadius.circular(20), color: Colors.white),
              child: Padding(
                padding: const EdgeInsets.symmetric(horizontal: 30),
                child: Column(
                  children: [
                    SizedBox(height: 30),
                    Text(
                      "Sign in using your existing account.",
                      style: TextStyle(fontSize: 15),
                      textAlign: TextAlign.center,
                    ),
                    SizedBox(height: 20),
                    textField(_emailController, 'Email', false),
                    // Whitespace

                    SizedBox(height: 10),

                    textField(_passwordController, 'Password', true),
                    // Whitespace
                    SizedBox(height: 15),

                    ElevatedButton(
                      child: Text(
                        "Sign In",
                        style: TextStyle(fontSize: 15, color: Colors.white),
                      ),
                      onPressed: () {
                        signIn(context);
                      },
                      style: ElevatedButton.styleFrom(
                          backgroundColor: Color(0xffDAB4DD),
                          shape: RoundedRectangleBorder(
                              borderRadius: BorderRadius.circular(20))),
                    ),
                    SizedBox(height: 10),
                    Row(
                      mainAxisAlignment: MainAxisAlignment.center,
                      children: [
                        Text(
                          "Dont have an account?",
                          style: TextStyle(color: Colors.black, fontSize: 15),
                        ),
                        SizedBox(width: 5),
                        GestureDetector(
                          onTap: () {
                            Navigator.push(
                              context,
                              MaterialPageRoute(builder: (context) => SignUp()),
                            );
                          },
                          child: Text(
                            "Sign Up",
                            style: TextStyle(
                              color: Color(0xff7E3185),
                              fontSize: 15,
                            ),
                          ),
                        )
                      ],
                    )
                  ],
                ),
              ),
            ),
          ],
        ),
      ),
    ));
  }

  void signIn(BuildContext context) async {
    final email = _emailController.text;
    final password = _passwordController.text;
    try {
      await FirebaseAuth.instance
          .signInWithEmailAndPassword(email: email, password: password);

      // Display the success message and wait until the user dismisses it
      await showMessage(context, "Login Successful!");

      // Navigate to the homepage after the dialog is dismissed
      Navigator.pushNamed(context, '/homepage');
    } on FirebaseAuthException catch (e) {
      if (e.code == 'user-not-found') {
        await showMessage(context, 'User not found!');
      } else if (e.code == 'wrong-password') {
        await showMessage(context, 'Incorrect Password!');
      }
    }
  }
}

Future<void> showMessage(BuildContext context, String message) async {
  return showDialog(
    context: context,
    builder: (context) {
      return AlertDialog(
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
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
