// ----------------------------------------------------------------------------
// TotalityHDR.js
// Composes manifest-defined eclipse exposure ladders with HDRComposition.
// ----------------------------------------------------------------------------

#feature-id    Utilities > TotalityHDR
#feature-info  Composes aligned Totality exposure blocks prepared by AsiToPix.

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

function validateBlock( block, manifestPath )
{
   if ( block === null || typeof block != "object" )
      throw new Error( "Invalid HDR block in manifest: " + manifestPath );
   if ( typeof block.blockNumber != "number" )
      throw new Error( "An HDR block has no numeric block number: " + manifestPath );
   if ( !(block.inputFiles instanceof Array) || block.inputFiles.length < 2 )
      throw new Error( "HDR block " + block.blockNumber + " must contain at least two input files." );
   if ( typeof block.outputPath != "string" || block.outputPath.length == 0 )
      throw new Error( "HDR block " + block.blockNumber + " has no output path." );

   for ( var i = 0; i < block.inputFiles.length; ++i )
      if ( typeof block.inputFiles[i] != "string" || !File.exists( block.inputFiles[i] ) )
         throw new Error( "HDR input frame not found: " + block.inputFiles[i] );

   var outputDirectory = File.extractDrive( block.outputPath ) +
                         File.extractDirectory( block.outputPath );
   if ( !File.directoryExists( outputDirectory ) )
      throw new Error( "HDR output directory not found: " + outputDirectory );
   if ( File.exists( block.outputPath ) )
      throw new Error( "HDR output already exists and will not be overwritten: " + block.outputPath );
}

function validateManifest( manifest, manifestPath )
{
   if ( manifest === null || typeof manifest != "object" )
      throw new Error( "Invalid HDR manifest: " + manifestPath );
   if ( manifest.schemaVersion != 2 )
      throw new Error( "Unsupported HDR manifest schema in: " + manifestPath );
   if ( typeof manifest.statusPath != "string" || manifest.statusPath.length == 0 )
      throw new Error( "The HDR manifest has no status path: " + manifestPath );
   if ( !(manifest.blocks instanceof Array) || manifest.blocks.length == 0 )
      throw new Error( "The HDR manifest has no blocks: " + manifestPath );

   for ( var i = 0; i < manifest.blocks.length; ++i )
      validateBlock( manifest.blocks[i], manifestPath );
}

function composeBlock( block )
{
   var P = ProcessInstance.fromIcon( "HDRComposition" );
   if ( P === null )
   {
      console.warningln( "<end><cbr>Process icon 'HDRComposition' was not found; using default HDRComposition settings." );
      P = new HDRComposition;
   }
   var images = [];
   for ( var i = 0; i < block.inputFiles.length; ++i )
      images.push( [ true, block.inputFiles[i] ] );
   P.images = images;

   console.writeln( "<end><cbr><br>Composing HDR block ", block.blockNumber,
                    " (", block.inputFiles.length, " exposures) to:<br><raw>",
                    block.outputPath, "</raw>" );
   if ( !P.executeGlobal() )
      throw new Error( "HDRComposition failed for block " + block.blockNumber +
                       ": " + block.outputPath );

   var outputWindow = ImageWindow.activeWindow;
   if ( outputWindow.isNull )
      throw new Error( "HDRComposition did not create an image window for block " +
                       block.blockNumber + "." );
   if ( File.exists( block.outputPath ) )
   {
      outputWindow.forceClose();
      throw new Error( "HDR output appeared during processing and will not be overwritten: " +
                       block.outputPath );
   }

   outputWindow.saveAs( block.outputPath, false, false, false, false );
   if ( !File.exists( block.outputPath ) )
   {
      outputWindow.forceClose();
      throw new Error( "Failed to save HDR output: " + block.outputPath );
   }
   outputWindow.forceClose();
   console.writeln( "<end><cbr>Saved HDR output:<br><raw>", block.outputPath, "</raw>" );
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
      currentBlockNumber: null,
      completedBlockNumbers: [],
      errorMessage: "",
      errorStack: ""
   };

   try
   {
      if ( !File.exists( manifestPath ) )
         throw new Error( "HDR manifest not found: " + manifestPath );

      var manifest = JSON.parse( File.readTextFile( manifestPath ) );
      if ( manifest !== null && typeof manifest == "object" &&
           typeof manifest.statusPath == "string" )
         statusPath = manifest.statusPath;
      writeStatus( statusPath, status );
      validateManifest( manifest, manifestPath );

      status.state = "running";
      writeStatus( statusPath, status );
      for ( var i = 0; i < manifest.blocks.length; ++i )
      {
         var block = manifest.blocks[i];
         status.currentBlockNumber = block.blockNumber;
         writeStatus( statusPath, status );
         composeBlock( block );
         status.completedBlockNumbers.push( block.blockNumber );
         writeStatus( statusPath, status );
      }

      status.state = "completed";
      status.currentBlockNumber = null;
      writeStatus( statusPath, status );
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
         console.criticalln( "<end><cbr>Unable to write HDR status: ",
                             errorText( statusError ) );
      }
      console.criticalln( "<end><cbr>TotalityHDR failed: ", status.errorMessage );
      throw error;
   }
}

main();
