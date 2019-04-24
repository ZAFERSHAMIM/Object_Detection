#define BLYNK_USE_DIRECT_CONNECT

// You could use a spare Hardware Serial on boards that have it (like Mega)
#include <SoftwareSerial.h>
SoftwareSerial DebugSerial(2, 3); // RX, TX

#define BLYNK_PRINT DebugSerial
#include <BlynkSimpleSerialBLE.h>
BlynkTimer timer;

// You should get Auth Token in the Blynk App.
// Go to the Project Settings (nut icon).
char auth[] = "YourAuthToken";

void notifyOnButtonPress()
{
  int isButtonPressed = digitalRead(D1);
  if (isButtonPressed==1 && flag==0) {
    Serial.println("Someone Opened the door");

    
    //   We allow 1 notification per 15 seconds for now.
    Blynk.notify("Alert : Someone Opened the door");
    flag=1;
  }
  else if (isButtonPressed==0)
  {
    flag=0;
  }
  
 
}
void setup()
{
  // Debug console
  Serial.begin(38400);

Blynk.begin(Serial, auth);
  // Setup notification button on pin D1 (the pin as per the ckt connection of door sensor,here D1 is of node mcu)
  pinMode(D1,INPUT_PULLUP);
 timer.setInterval(16000L,notifyOnButtonPress);
 
}
void loop()
{
  Blynk.run();
  timer.run(); // Initiates BlynkTimer
}
