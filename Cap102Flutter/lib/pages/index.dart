// ignore_for_file: prefer_const_constructors

import 'package:cap102/pages/signUp.dart';
import 'package:flutter/material.dart';

// ignore: camel_case_types
class IntroPage extends StatelessWidget {
  const IntroPage({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Color(0xffDAB4DD),
      body: Center(
        child: Padding(
          padding: const EdgeInsets.symmetric(vertical: 50.0),
          child: Column(
            children: [
              Text(
                "Portalbe Printing for Everyone",
                style: TextStyle(
                    color: Colors.white,
                    fontSize: 50,
                    fontWeight: FontWeight.w500),
                textAlign: TextAlign.center,
              ),

              // Spacing
              const SizedBox(height: 200),

              // Brand
              Text("PrintVendo",
                  style: TextStyle(
                    color: Colors.white,
                    fontSize: 30,
                  )),

              // Spacing
              const SizedBox(height: 250),

              //Get Started Button

              ElevatedButton(
                child: Text(
                  "Get Started",
                  style: TextStyle(color: Color(0xff7E3185)),
                ),
                onPressed: () {
                  // navigate to Sing Up Page
                  Navigator.push(context,
                      MaterialPageRoute(builder: (context) => SignUp()));
                },
              )
            ],
          ),
        ),
      ),
    );
  }
}
