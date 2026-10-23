const {defineConfig}=require('@playwright/test');
module.exports=defineConfig({testDir:'tests/ui',timeout:25000,workers:1,use:{baseURL:'http://127.0.0.1:8765',viewport:{width:1672,height:941},headless:true,timezoneId:'Europe/Kyiv'},webServer:{command:'python -m http.server 8765 --bind 127.0.0.1',url:'http://127.0.0.1:8765/taskpane.html',reuseExistingServer:true},reporter:'list'});
