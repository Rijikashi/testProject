# Sideload

https://mail.novacoast.com
* Options > General >  Manage Plugins > +
* Add from file
* Upload manifest.xml from this directory


# Setup NPM
```
npm install
```

# Build App
```
node node_modules/webpack/bin/webpack.js  --config webpack.config.js --mode development
```

# Run Server
```
python python-server/simple-https-server.py 
```

# Enable "Inspect Element" in the taskpane:
run this command in terminal:
```
defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true
```

# Clearing Plugin Cache in MAC
If changes are not appearing in Outlook, delete the following files to completely reinstall the plugin from manifest:
```
~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/
~/Library/Containers/com.microsoft.Outlook/Data/Library/Application Support/Microsoft/Office/16.0/Wef/
```