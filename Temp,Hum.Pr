#define BLYNK_TEMPLATE_ID "TMPL3V3E93rdw"
#define BLYNK_TEMPLATE_NAME "Temp IOT.Pr"
#define BLYNK_AUTH_TOKEN "Your Auth Token"
#include <WiFi.h>
#include <Blynk Simple Esp32.h>
#include <DHT-h>
char SSID[]="WiFi SSID";
char Pass[]="Wifi Pass";
#define DHTPIN 4
#define DHTTYPE DHT 11
DHT dht(DHTPIN,DHTTYPE);
void setup()
{
serial.begin(115200);
dht.begin();
Blynk.begin(BLYNK_AUTH_TOKEN,SSID,Pass);
}
void loop()
{
float t=dht.read Temperature();
float h=dht.read Humidity();
Blynk.virtual write(v0,t);
Blynk.virtual write(v1,h);
delay(2000);
Blynk.run();
}
