// ----------------------------------------------------------------------------
// AlignTotality.js
// Applies RGB channel offsets captured from a ChannelMatch process icon, then
// applies optional orthogonal FastRotation transforms to integrated frames.
// ----------------------------------------------------------------------------

#feature-id    Utilities > AlignTotality
#feature-info  Channel-aligns and optionally rotates or mirrors integrated Totality blocks prepared by AsiToPix.

// ASITOPIX_IPC_MANIFEST

function argumentValue( name )
{
   if ( name == "manifest" && typeof ASITOPIX_MANIFEST_PATH != "undefined" )
      return ASITOPIX_MANIFEST_PATH;

   var prefix = name + "=";
   for ( var i = 0; i < jsArguments.length; ++i )
      if ( jsArguments[i].indexOf( prefix ) == 0 )
         return jsArguments[i].substring( prefix.length );
   throw new Error( "Missing required command-line argument: " + name );
}

function errorText( error )
{
   if ( error === null || error === undefined )
      return "Unknown PJSR error.";
   if ( typeof error.message == "string" && error.message.length > 0 )
      return error.message;
   return error.toString();
}

function writeStatus( statusPath, status )
{
   if ( typeof statusPath == "string" && statusPath.length > 0 )
      File.writeTextFile( statusPath, JSON.stringify( status, null, 2 ) );
}

function validateChannelObjects( channels, context )
{
   if ( !(channels instanceof Array) || channels.length != 3 )
      throw new Error( context + " must contain exactly three RGB channels." );

   var result = [];
   var names = [ "R", "G", "B" ];
   for ( var i = 0; i < 3; ++i )
   {
      var channel = channels[i];
      if ( channel === null || typeof channel != "object" )
         throw new Error( "Invalid " + names[i] + " channel in " + context + "." );
      var dx = Number( channel.dx );
      var dy = Number( channel.dy );
      if ( !isFinite( dx ) || !isFinite( dy ) )
         throw new Error( "Non-finite " + names[i] + " offset in " + context + "." );
      result.push( {
         name: names[i],
         enabled: Boolean( channel.enabled ),
         dx: dx,
         dy: dy
      } );
   }
   return result;
}

function channelObjectsFromProcess( process, iconId )
{
   if ( process === null || !(process instanceof ChannelMatch) )
      throw new Error( "Process icon '" + iconId + "' is not a ChannelMatch instance." );
   if ( !(process.channels instanceof Array) || process.channels.length != 3 )
      throw new Error( "ChannelMatch icon '" + iconId + "' has an invalid channels table." );

   var names = [ "R", "G", "B" ];
   var result = [];
   for ( var i = 0; i < 3; ++i )
   {
      var row = process.channels[i];
      if ( !(row instanceof Array) || row.length < 4 )
         throw new Error( "ChannelMatch icon '" + iconId + "' has an invalid " + names[i] + " row." );
      var dx = Number( row[1] );
      var dy = Number( row[2] );
      if ( !isFinite( dx ) || !isFinite( dy ) )
         throw new Error( "ChannelMatch icon '" + iconId + "' has non-finite " + names[i] + " offsets." );
      result.push( {
         name: names[i],
         enabled: Boolean( row[0] ),
         dx: dx,
         dy: dy
      } );
   }
   return result;
}

function channelRows( channels )
{
   var rows = [];
   for ( var i = 0; i < channels.length; ++i )
      rows.push( [ channels[i].enabled, channels[i].dx, channels[i].dy, 1.0 ] );
   return rows;
}

function findNewChannelMatchIcon( baselineIconIds )
{
   var baseline = {};
   for ( var i = 0; i < baselineIconIds.length; ++i )
      baseline[baselineIconIds[i]] = true;

   var current = ProcessInstance.iconsByProcessId( "ChannelMatch" );
   var added = [];
   for ( var j = 0; j < current.length; ++j )
      if ( !baseline[current[j]] )
         added.push( current[j] );

   if ( added.length == 0 )
      throw new Error(
         "No new ChannelMatch process icon was found. After adjusting ChannelMatch, " +
         "drag its blue New Instance triangle to an empty place in the PixInsight workspace." );
   if ( added.length > 1 )
      throw new Error(
         "More than one new ChannelMatch process icon was found: " + added.join( ", " ) +
         ". Keep one newly created ChannelMatch icon and try again." );
   return added[0];
}

function saveCapturedSettings( path, iconId, channels )
{
   var settings = {
      schemaVersion: 1,
      sourceIconId: iconId,
      capturedAtUtc: (new Date).toISOString(),
      channels: channels
   };
   File.writeTextFile( path, JSON.stringify( settings, null, 2 ) + "\n" );
   if ( !File.exists( path ) )
      throw new Error( "Failed to save ChannelMatch settings: " + path );
   return settings;
}

function validateTransform( transform )
{
   if ( transform === null || typeof transform != "object" )
      throw new Error( "Missing Totality alignment transform." );
   var rotation = Number( transform.rotation );
   if ( rotation != 0 && rotation != 90 && rotation != 180 && rotation != 270 )
      throw new Error( "Unsupported Totality alignment rotation: " + transform.rotation );
   if ( transform.flip != "None" &&
        transform.flip != "Horizontal" && transform.flip != "Vertical" )
      throw new Error( "Unsupported Totality alignment mirror: " + transform.flip );
   return { rotation: rotation, flip: transform.flip };
}

function validateManifest( manifest, manifestPath )
{
   if ( manifest === null || typeof manifest != "object" )
      throw new Error( "Invalid Totality alignment manifest: " + manifestPath );
   if ( manifest.schemaVersion != 1 )
      throw new Error( "Unsupported Totality alignment manifest schema: " + manifest.schemaVersion );
   if ( manifest.operation != "inspect" && manifest.operation != "align" )
      throw new Error( "Unsupported Totality alignment operation: " + manifest.operation );
   if ( typeof manifest.statusPath != "string" || manifest.statusPath.length == 0 )
      throw new Error( "Missing Totality alignment status path." );

   if ( manifest.operation == "inspect" )
      return;
   if ( manifest.channelMatchSource != "saved" && manifest.channelMatchSource != "capture" )
      throw new Error( "Unsupported ChannelMatch source: " + manifest.channelMatchSource );
   if ( manifest.channelMatchSource == "saved" )
      validateChannelObjects( manifest.channels, "saved ChannelMatch settings" );
   else
   {
      if ( !(manifest.baselineIconIds instanceof Array) )
         throw new Error( "Missing baseline ChannelMatch process icon list." );
      if ( typeof manifest.settingsPath != "string" || manifest.settingsPath.length == 0 )
         throw new Error( "Missing ChannelMatch settings output path." );
   }
   validateTransform( manifest.transform );
   if ( !(manifest.frames instanceof Array) || manifest.frames.length == 0 )
      throw new Error( "The Totality alignment manifest has no frames." );
   for ( var i = 0; i < manifest.frames.length; ++i )
   {
      var frame = manifest.frames[i];
      if ( frame === null || typeof frame != "object" ||
           typeof frame.inputPath != "string" || frame.inputPath.length == 0 ||
           typeof frame.outputPath != "string" || frame.outputPath.length == 0 )
         throw new Error( "Invalid Totality alignment frame at manifest index " + i + "." );
      if ( !File.exists( frame.inputPath ) )
         throw new Error( "Totality alignment input frame not found: " + frame.inputPath );
      if ( File.exists( frame.outputPath ) )
         throw new Error( "Totality alignment output already exists and will not be overwritten: " + frame.outputPath );
   }
}

function openSingleImage( path )
{
   var windows = ImageWindow.open( path );
   if ( windows.length != 1 )
   {
      for ( var i = 0; i < windows.length; ++i )
         if ( windows[i] !== null && !windows[i].isNull )
            windows[i].forceClose();
      throw new Error( "Expected one image in XISF file, got " + windows.length + ": " + path );
   }
   return windows[0];
}

function applyChannelMatch( view, channels )
{
   var process = new ChannelMatch;
   process.channels = channelRows( channels );
   if ( !process.executeOn( view ) )
      throw new Error( "ChannelMatch execution failed on view: " + view.fullId );
}

function applyFastRotation( view, mode )
{
   var process = new FastRotation;
   process.mode = mode;
   process.noGUIMessages = true;
   if ( !process.executeOn( view ) )
      throw new Error( "FastRotation execution failed on view: " + view.fullId );
}

function applyTransforms( view, transform )
{
   if ( transform.rotation == 90 )
      applyFastRotation( view, FastRotation.prototype.Rotate90CW );
   else if ( transform.rotation == 180 )
      applyFastRotation( view, FastRotation.prototype.Rotate180 );
   else if ( transform.rotation == 270 )
      applyFastRotation( view, FastRotation.prototype.Rotate90CCW );

   if ( transform.flip == "Horizontal" )
      applyFastRotation( view, FastRotation.prototype.HorizontalMirror );
   else if ( transform.flip == "Vertical" )
      applyFastRotation( view, FastRotation.prototype.VerticalMirror );
}

function processFrame( frame, channels, transform )
{
   var window = null;
   try
   {
      window = openSingleImage( frame.inputPath );
      applyChannelMatch( window.mainView, channels );
      applyTransforms( window.mainView, transform );
      window.saveAs( frame.outputPath, false, false, false, false );
      if ( !File.exists( frame.outputPath ) )
         throw new Error( "Failed to save aligned Totality frame: " + frame.outputPath );
   }
   finally
   {
      if ( window !== null && !window.isNull )
         window.forceClose();
   }
}

function main()
{
   console.show();
   console.abortEnabled = true;

   var manifestPath = argumentValue( "manifest" );
   var statusPath = "";
   var status = {
      schemaVersion: 1,
      state: "starting",
      currentInputPath: "",
      completedOutputPaths: [],
      channelMatchIcons: [],
      channelMatch: null,
      errorMessage: "",
      errorStack: ""
   };

   try
   {
      if ( !File.exists( manifestPath ) )
         throw new Error( "Totality alignment manifest not found: " + manifestPath );
      var manifest = JSON.parse( File.readTextFile( manifestPath ) );
      if ( manifest !== null && typeof manifest == "object" &&
           typeof manifest.statusPath == "string" )
         statusPath = manifest.statusPath;
      writeStatus( statusPath, status );
      validateManifest( manifest, manifestPath );

      if ( manifest.operation == "inspect" )
      {
         status.channelMatchIcons = ProcessInstance.iconsByProcessId( "ChannelMatch" );
         status.state = "completed";
         writeStatus( statusPath, status );
         return;
      }

      var channels;
      var sourceIconId = "";
      if ( manifest.channelMatchSource == "saved" )
         channels = validateChannelObjects( manifest.channels, "saved ChannelMatch settings" );
      else
      {
         sourceIconId = findNewChannelMatchIcon( manifest.baselineIconIds );
         channels = channelObjectsFromProcess(
            ProcessInstance.fromIcon( sourceIconId ), sourceIconId );
         status.channelMatch = saveCapturedSettings(
            manifest.settingsPath, sourceIconId, channels );
      }
      if ( status.channelMatch === null )
         status.channelMatch = {
            schemaVersion: 1,
            sourceIconId: "saved JSON",
            capturedAtUtc: "",
            channels: channels
         };

      var transform = validateTransform( manifest.transform );
      status.state = "running";
      writeStatus( statusPath, status );
      for ( var i = 0; i < manifest.frames.length; ++i )
      {
         var frame = manifest.frames[i];
         status.currentInputPath = frame.inputPath;
         writeStatus( statusPath, status );
         processFrame( frame, channels, transform );
         status.completedOutputPaths.push( frame.outputPath );
         writeStatus( statusPath, status );
      }

      status.state = "completed";
      status.currentInputPath = "";
      writeStatus( statusPath, status );
      console.writeln( "<end><cbr><br>Aligned ", manifest.frames.length,
                       " Totality frame(s)." );
   }
   catch ( error )
   {
      status.state = "failed";
      status.errorMessage = errorText( error );
      status.errorStack = error !== null && error !== undefined &&
                          typeof error.stack == "string" ? error.stack : "";
      try
      {
         writeStatus( statusPath, status );
      }
      catch ( statusError )
      {
         console.criticalln( "<end><cbr>Unable to write alignment status: ",
                             errorText( statusError ) );
      }
      console.criticalln( "<end><cbr>AlignTotality failed: ", status.errorMessage );
      throw error;
   }
}

main();
