# PCF - Attachments Grid

A custom PCF control developed by [Josh Hetherington](https://github.com/jhetheringt7) and [Ben Bartle](https://github.com/benlbartle) using the [PowerApps Component Framework](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/overview). This is designed to offer a better UX than the standard notes pane for viewing attachments.

You can simple drag and drop one or multiple files into the attachments grid and they'll automatically be added to the notes.

## Pre-requisites 

In order to build and deploy these to your CDS instance you'll need the following:

- [NodeJS & npm](https://nodejs.org/en/)
- [The PCF CLI](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/create-custom-controls-using-pcf)
- Either:
    - [Visual Studio Code](https://code.visualstudio.com/) (with [.NET Core SDK](https://dotnet.microsoft.com/download))
    - [Visual Studio](https://visualstudio.microsoft.com/) (2017 or later)

If you want to deploy this and test it you'll obviously need some sort of PowerApps/Dynamics 365 license and instance. Grab a trial from [here](https://trials.dynamics.com/).

## Building & Debugging

Once you've grabbed the code, navigate to the correct folder for the component you want, and run the following command from the terminal:

```shell
npm install
```

This will install all the dependencies (and make take a minute or two). Then run:

```shell
npm run start
```

This should bootstrap the component and run the harness to allow you to see the component running and debug if required.

### Usage

Simply add any field to the Form, and configure it to use the control. It doesn't really matter what the field is as the control only uses the supplied `ComponentFramework.Context`to grab the entity id and entity logical name. Your entity will need to support attachments (obviously) for this to work.

Then drag your files onto the control and they'll be uploaded and the grid will refresh.

### Contributing

This is still very much a work in progress, and we would like to investigate the following:

1. Better CSS Styling, if anyone has any recommendations, interested in working with [Microsoft's Office Fabric](https://developer.microsoft.com/en-us/fabric#/) when it becomes more available
2. Better error handling for loading the control on an un-saved record, and for records which don't support notes
3. Tests!

Just fork the repo, create a feature branch for your change and send us a PR.