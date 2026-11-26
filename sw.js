/* Republica | 1963 · Calendar — fișier de fundal minim, ca aplicația să se poată instala.
   Nu păstrează nimic în cache: datele vin mereu proaspete din Sheet. */
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", event => event.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {});
