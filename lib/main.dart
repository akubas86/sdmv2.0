import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path/path.dart' as p;
import 'package:audioplayers/audioplayers.dart';


void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Sistem Daftar Makmal',
      home: const Sdm(),
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        scaffoldBackgroundColor: Colors.blueGrey[100],
      ),
    );
  }
}

class _PressableScale extends StatefulWidget {
  final Widget child;
  final VoidCallback onTap;

  const _PressableScale({
    required this.child,
    required this.onTap,
  });

  @override
  State<_PressableScale> createState() => _PressableScaleState();
}

class _PressableScaleState extends State<_PressableScale>
    with SingleTickerProviderStateMixin {

  late AnimationController _controller;
  final AudioPlayer _player = AudioPlayer();

  @override
  void initState() {
    super.initState();
    _controller = AnimationController(
      vsync: this,
      duration: Duration(milliseconds: 80),
      lowerBound: 0.0,
      upperBound: 0.05, // amount to shrink
    );

    _player.setVolume(0.5);     // softer click
  }

  @override
  void dispose() {
    _controller.dispose();
    super.dispose();
  }

  void _onTapDown(TapDownDetails details) {
    _controller.forward(); // shrink down
  }

  void _onTapUp(TapUpDetails details) async {
    _controller.reverse(); // bounce back

    // play click sound
    await _player.play(AssetSource('sfx/click.wav'));

    widget.onTap();
  }

  void _onTapCancel() {
    _controller.reverse();
  }

  @override
  Widget build(BuildContext context) {
    return GestureDetector(
      onTapDown: _onTapDown,
      onTapUp: _onTapUp,
      onTapCancel: _onTapCancel,
      child: AnimatedBuilder(
        animation: _controller,
        builder: (context, child) {
          double scale = 1 - _controller.value;
          return Transform.scale(
            scale: scale,
            child: child,
          );
        },
        child: widget.child,
      ),
    );
  }
}


class Sdm extends StatelessWidget {
  const Sdm({super.key});

  Widget buildMenuButton({
    required Widget icon,
    required String label,
    required VoidCallback onPressed,
  }) {
    return _PressableScale(
      onTap: onPressed,
      child: SizedBox(
        width: 180,
        height: 180,
        child: Container(
          padding: EdgeInsets.all(20),
          decoration: BoxDecoration(
            color: Colors.lightBlueAccent,
            borderRadius: BorderRadius.circular(20),
            boxShadow: [
              BoxShadow(
                color: Colors.black26,
                blurRadius: 10,
                offset: Offset(0, 6),
              ),
            ],
          ),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.spaceEvenly,
            children: [
              icon,
              SizedBox(height: 10),
              Text(
                label,
                style: TextStyle(fontSize: 16, fontWeight: FontWeight.w600),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> runDaftarProgram() async {
    bool isRelease = bool.fromEnvironment('dart.vm.product');

    // Use fixed install path in release, dynamic in debug
    final appDir = isRelease
        ? r'C:\Program Files (x86)\sdm'
        : r'C:\Users\USER\Desktop\sdm';  // dev path

    final pythonExe = p.join(appDir, 'python-script', 'BD_0.0.13.exe');

    final result = await Process.run(
      pythonExe,
      [],
    );

    print('stdout: ${result.stdout}');
    print('stderr: ${result.stderr}');

  }

  Future<void> runKemaskiniProgram() async {
    bool isRelease = bool.fromEnvironment('dart.vm.product');

    // Use fixed install path in release, dynamic in debug
    final appDir = isRelease
        ? r'C:\Program Files (x86)\sdm'
        : r'C:\Users\USER\Desktop\sdm';  // dev path

    final pythonExe = p.join(appDir, 'python-script', 'KK_0.0.5.exe');

    final result = await Process.run(
      pythonExe,
      [],
    );

    print('stdout: ${result.stdout}');
    print('stderr: ${result.stderr}');

  }

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Colors.blueGrey[100],
        title: Align(
          alignment: Alignment.centerRight,
          child: Text('SDM V2.0',
          style: TextStyle(
            fontSize: 12,
          ),)),
      ),
      body: Padding(
          padding: const EdgeInsets.all(16.0),
          child: Column(
            children: [
              Text("Sistem Daftar Makmal",
              style: TextStyle(
                fontSize: 48,
                fontWeight: FontWeight.bold,
                //color: Colors.blue,
              ),
                textAlign: TextAlign.center,
              ),
              SizedBox(height: 30),
              Center(
                child: Column(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    Row(
                      mainAxisAlignment: MainAxisAlignment.center,
                      children: [
                        Padding(
                          padding: const EdgeInsets.all(4.0),
                          child: buildMenuButton(
                              icon: Image.asset(
                                "assets/icon/dbicon.png",
                                width: 80,
                                height: 80,
                              ),
                              label: "Buku Daftar",
                              onPressed: () => runDaftarProgram(),
                          ),
                        ),
                        const SizedBox(width: 16),
                        Padding(
                          padding: const EdgeInsets.all(4.0),
                          child: buildMenuButton(
                            icon: Image.asset(
                              "assets/icon/kkicon.png",
                              width: 80,
                              height: 80,
                            ),
                            label: "Kemaskini",
                            onPressed: () => runKemaskiniProgram(),
                          ),
                        )
                      ],
                    ),
                    const SizedBox(width: 16),
                    Row(
                      mainAxisAlignment: MainAxisAlignment.center,
                      children: [
                          Padding(
                            padding: const EdgeInsets.all(4.0),
                            child: buildMenuButton(
                                icon: Image.asset(
                                  "assets/icon/ppicon.png",
                                  width: 80,
                                  height: 80,
                                ),
                                label: "Piagam Pelanggan",
                                onPressed: () {}
                            ),
                          ),
                          const SizedBox(width: 16),
                          Padding(
                            padding: const EdgeInsets.all(4.0),
                            child: buildMenuButton(
                                icon: Image.asset(
                                  "assets/icon/ppicon.png",
                                  width: 80,
                                  height: 80,
                                ),
                                label: "Dashboard",
                                onPressed: () {}
                            ),
                          ),
                      ],
                    ),

                  ],
                ),
              ),
            ],
          ),
      ),
    );

  }
}




