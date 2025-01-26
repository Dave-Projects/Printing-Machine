// ignore_for_file: use_build_context_synchronously

import 'package:cap102/pages/signIn.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:flutter/material.dart';
import 'package:fluttertoast/fluttertoast.dart';

class AuthPage extends StatefulWidget {
  const AuthPage({super.key});

  @override
  State<AuthPage> createState() => _AuthPageState();
}

class _AuthPageState extends State<AuthPage> {
  @override
  Widget build(BuildContext context) {
    return const Scaffold();
  }
}

class AuthService {
  Future<void> signUp(
      {required String email,
      required String password,
      required BuildContext context}) async {
    try {
      await FirebaseAuth.instance
          .createUserWithEmailAndPassword(email: email, password: password);
      Navigator.pushNamed(context, '/signIn');
    } on FirebaseAuthException catch (e) {
      if (e.code == 'weak-password') {
        showErrorMessage(context, "Password length is too short.");
      } else if (e.code == 'email-already-in-use') {
        showErrorMessage(context, "Email already exists.");
      }
      // ignore: empty_catches
    } catch (e) {}
  }

  Future<void> signIn({
    required String email,
    required String password,
    required BuildContext context,
  }) async {
    // Show the loading indicator

    showDialog(
      context: context,
      builder: (context) {
        return const Center(
          child: CircularProgressIndicator(),
        );
      },
    );

    try {
      await FirebaseAuth.instance
          .signInWithEmailAndPassword(email: email, password: password);

      //  Navigator.pop(context);
      Navigator.pushNamed(context, '/homepage');
    } on FirebaseAuthException catch (e) {
      // Close the loading indicator
      //Navigator.pop(context);

      if (e.code == 'user-not-found' || e.code == 'wrong-password') {
        Fluttertoast.showToast(msg: e.code, gravity: ToastGravity.CENTER);
      }
    }
  }

  Future<void> signOut({required BuildContext context}) async {
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
        builder: (BuildContext context) => const SignIn(),
      ),
    );
  }

  void showErrorMessage(BuildContext context, String message) {
    showDialog(
      context: context,
      builder: (context) {
        return AlertDialog(
          title: Center(
            child: Text(
              message,
              style: const TextStyle(color: Colors.black, fontSize: 20),
            ),
          ),
          actions: [
            MaterialButton(
              onPressed: () {
                Navigator.pop(context);
              },
              child: const Text("Okay"),
            ),
          ],
        );
      },
    );
  }
}
