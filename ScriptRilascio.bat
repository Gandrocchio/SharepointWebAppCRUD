call gulp clean
call gulp bundle --ship
call gulp package-solution --ship
ren .\temp\deploy\appcrud-helloworldwebpartstrings*.js       appcrud-helloworldwebpartstrings_en-us_536e65149b0acf4d52c0043073b9fc59.js
ren .\temp\deploy\hello-world-web-part_*.js     hello-world-web-part_b88e9baa4ef2dfe4850406642231d115.js
