#define HP1sensorPin 3
#define HP1relayPin 4 
#define HP2sensorPin 5
#define HP2relayPin 6

#define coinSlot 9
int coinSlotStatus;
int coinReceived; //pulses

int pulseout1;
int targetPulses1 = 0;
boolean relayActive1 = false;
int sensorState1 = 0;
int lastSensorState1 = 0;

int pulseout2;
int targetPulses2 = 0;
boolean relayActive2 = false;
int sensorState2 = 0;
int lastSensorState2 = 0;

bool hopper1Dispensing = false;
bool hopper5Dispensing = false;
String Status = "idle";

void setup() {
  Serial.begin(9600);
  pinMode(coinSlot, INPUT_PULLUP);

  // Hopper 1 setup
  pinMode(HP1sensorPin, INPUT_PULLUP);
  pinMode(HP1relayPin, OUTPUT);
  digitalWrite(HP1relayPin, LOW); // Assuming relay is active LOW

  // Hopper 2 setup
  pinMode(HP2sensorPin, INPUT_PULLUP);
  pinMode(HP2relayPin, OUTPUT);
  digitalWrite(HP2relayPin, LOW); // Assuming relay is active LOW
}
void checkStatus(){
  if (hopper1Dispensing == true || hopper5Dispensing == true){
    Status = "dispensing";
  }else{
    Status = "idle";
  }
}

void loop() {
  checkStatus();

  // Hopper 1 logic
  sensorState1 = digitalRead(HP1sensorPin);

  if (Serial.available() > 0) {
    // Read the incoming data
    String command = Serial.readStringUntil('\n');
    if (command.startsWith("Change5Peso: ")) {
      targetPulses1 = command.substring(13).toInt();
      pulseout1 = 0;
      relayActive1 = true;
      digitalWrite(HP1relayPin, HIGH);
      hopper5Dispensing = true;
      // Serial.print("Hopper 1 Target pulses set to ");
      // Serial.println(targetPulses1);
      checkStatus();
      if (Status == "dispensing"){
        int remainingCoins = targetPulses1 - pulseout1;
        String message = "D5: " + String(remainingCoins);
        Serial.println(message);
      }else{
        Serial.println(Status);
      }
      delay(130);
    } else if (command.startsWith("Change1Peso: ")) {
      targetPulses2 = command.substring(13).toInt();
      pulseout2 = 0;
      relayActive2 = true;
      digitalWrite(HP2relayPin, HIGH);
      hopper1Dispensing = true;
      // Serial.print("Hopper 2 Target pulses set to ");
      // Serial.println(targetPulses2);
      checkStatus();
      if (Status == "dispensing"){
        int remainingCoins = targetPulses2 - pulseout2;
        String message = "D1: " + String(remainingCoins);
        Serial.println(message);
      }else{
        Serial.println(Status);
      }
      delay(5);

    }

    if (command == "Withdraw1") {
      while (Serial.available() == 0 || Serial.readStringUntil('\n') != "WithdrawStop") {
      digitalWrite(HP2relayPin, HIGH);
      }
      digitalWrite(HP2relayPin, LOW);
    }else if (command == "Withdraw5") {
      while (Serial.available() == 0 || Serial.readStringUntil('\n') != "WithdrawStop") {
      digitalWrite(HP1relayPin, HIGH);
      }
      digitalWrite(HP1relayPin, LOW);
    }

    // FOR COIN RECEIVING
    if (command == "reset") {
      coinReceived = 0;
    }

    // RESET Hoppers //Stop All
    if (command == "SA") { 
      digitalWrite(HP1relayPin, LOW);
      digitalWrite(HP2relayPin, LOW);
      relayActive1 = false;
      relayActive2 = false;

      Status = "idle";

      pulseout1 = 0;
      targetPulses1 = 0;
      hopper5Dispensing = false;

      pulseout2 = 0;
      targetPulses2 = 0;
      hopper1Dispensing = false;
    }
  }
  // Hopper 2 logic
  // 5 PESO
  if (sensorState1 == HIGH && lastSensorState1 == LOW) {
    if (relayActive1 && sensorState1 == HIGH) {
      pulseout1 += 1;
      // Serial.println("Hopper 1: " + String(pulseout1));

      if (pulseout1 >= targetPulses1) {
        digitalWrite(HP1relayPin, LOW);
        relayActive1 = false;
        pulseout1 = 0;
        targetPulses1 = 0;
        // Serial.println("Hopper 1: Relay stopped");
        hopper5Dispensing = false;
      }
      checkStatus();
      if (Status == "dispensing"){
        int remainingCoins = targetPulses1 - pulseout1;
        String message = "D5: " + String(remainingCoins);
        Serial.println(message);
      }else{
        Serial.println(Status);
      }
      delay(3);
    }
  }
  lastSensorState1 = sensorState1;

  // Hopper 2 logic
  // 1 PESO
  sensorState2 = digitalRead(HP2sensorPin);
  if (sensorState2 == HIGH && lastSensorState2 == LOW) {
    if (relayActive2 && sensorState2 == HIGH) {
      pulseout2 += 1;
      // Serial.println("Hopper 2: " + String(pulseout2));
      if (pulseout2 >= targetPulses2) {
        digitalWrite(HP2relayPin, LOW);
        relayActive2 = false;
        pulseout2 = 0;
        targetPulses2 = 0;
        // SerChanial.println("Hopper 2: Relay stopped");
        hopper1Dispensing = false;
      }
      checkStatus();
      if (Status == "dispensing"){
        int remainingCoins = targetPulses2 - pulseout2;
        String message = "D1: " + String(remainingCoins);
        Serial.println(message);
      }else{
        Serial.println(Status);
      }
      delay(3);
    }
  }
  lastSensorState2 = sensorState2;

  delay(3);

  // FOR COIN RECEIVING
  coinSlotStatus = digitalRead(coinSlot);
  if (coinSlotStatus == 0 ){
    coinReceived += 1;
    Serial.println("Coin: " + String(coinReceived));
    delay(100);
  }
}
