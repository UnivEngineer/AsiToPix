// ----------------------------------------------------------------------------
// IntegrateTotality.js
// Integrates manifest-defined eclipse frame blocks with ImageIntegration.
// ----------------------------------------------------------------------------

#feature-id    Utilities > IntegrateTotality
#feature-info  Integrates registered eclipse frame blocks prepared by AsiToPix.

function argumentValue( name )
{
   var prefix = name + "=";
   for ( var i = 0; i < jsArguments.length; ++i )
      if ( jsArguments[i].indexOf( prefix ) == 0 )
         return jsArguments[i].substring( prefix.length );

   if ( name == "manifest" )
   {
      var scriptFilePath = #__FILE__;
      var siblingManifestPath = File.extractDrive( scriptFilePath ) +
                                File.extractDirectory( scriptFilePath ) +
                                "/IntegrationPlan.json";
      if ( File.exists( siblingManifestPath ) )
         return siblingManifestPath;
   }

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
      throw new Error( "Invalid integration block in manifest: " + manifestPath );
   if ( typeof block.blockNumber != "number" )
      throw new Error( "An integration block has no numeric block number: " + manifestPath );
   if ( !(block.inputFiles instanceof Array) || block.inputFiles.length < 2 )
      throw new Error( "Integration block " + block.blockNumber + " must contain at least two input files." );
   if ( typeof block.outputPath != "string" || block.outputPath.length == 0 )
      throw new Error( "Integration block " + block.blockNumber + " has no output path." );

   for ( var i = 0; i < block.inputFiles.length; ++i )
      if ( typeof block.inputFiles[i] != "string" || !File.exists( block.inputFiles[i] ) )
         throw new Error( "Integration input frame not found: " + block.inputFiles[i] );

   var outputDirectory = File.extractDrive( block.outputPath ) +
                         File.extractDirectory( block.outputPath );
   if ( !File.directoryExists( outputDirectory ) )
      throw new Error( "Integration output directory not found: " + outputDirectory );
   if ( File.exists( block.outputPath ) )
      throw new Error( "Integration output already exists and will not be overwritten: " + block.outputPath );
}

function validateManifest( manifest, manifestPath )
{
   if ( manifest === null || typeof manifest != "object" )
      throw new Error( "Invalid integration manifest: " + manifestPath );
   if ( manifest.schemaVersion != 2 )
      throw new Error( "Unsupported integration manifest schema in: " + manifestPath );
   if ( typeof manifest.statusPath != "string" || manifest.statusPath.length == 0 )
      throw new Error( "The integration manifest has no status path: " + manifestPath );
   if ( !(manifest.blocks instanceof Array) || manifest.blocks.length == 0 )
      throw new Error( "The integration manifest has no blocks: " + manifestPath );

   for ( var i = 0; i < manifest.blocks.length; ++i )
      validateBlock( manifest.blocks[i], manifestPath );
}

function closeWindowById( id )
{
   if ( typeof id == "string" && id.length > 0 )
   {
      var window = ImageWindow.windowById( id );
      if ( !window.isNull )
         window.forceClose();
   }
}

function integrateBlock( block )
{
   var P = new ImageIntegration;
   var images = [];
   for ( var i = 0; i < block.inputFiles.length; ++i )
      images.push( [ true, block.inputFiles[i], "", "" ] );
   P.images = images;
   if ( P.images.length != block.inputFiles.length )
      throw new Error( "Failed to assign ImageIntegration source images for block " +
                       block.blockNumber + ": expected " + block.inputFiles.length +
                       ", assigned " + P.images.length + "." );
   P.inputHints = "fits-keywords normalize only-first-image raw cfa signed-is-physical";
   P.overrideImageType = false;
   P.imageType = 0;
   P.combination = ImageIntegration.prototype.Average;
   P.weightMode = ImageIntegration.prototype.DontCare;
   P.weightKeyword = "";
   P.csvWeightsFilePath = "";
   P.weightScale = ImageIntegration.prototype.WeightScale_BWMV;
   P.minWeight = 0.005000;
   P.csvWeights = "";
   P.adaptiveGridSize = 16;
   P.adaptiveNoScale = false;
   P.ignoreNoiseKeywords = false;
   P.normalization = ImageIntegration.prototype.NoNormalization;
   P.rejection = ImageIntegration.prototype.WinsorizedSigmaClip;
   P.rejectionNormalization = ImageIntegration.prototype.Scale;
   P.minMaxLow = 1;
   P.minMaxHigh = 1;
   P.pcClipLow = 0.200;
   P.pcClipHigh = 0.100;
   P.sigmaLow = 3.500;
   P.sigmaHigh = 4.000;
   P.winsorizationCutoff = 5.000;
   P.linearFitLow = 5.000;
   P.linearFitHigh = 4.000;
   P.esdOutliersFraction = 0.30;
   P.esdAlpha = 0.05;
   P.esdLowRelaxation = 1.00;
   P.rcrLimit = 0.10;
   P.ccdGain = 1.00;
   P.ccdReadNoise = 10.00;
   P.ccdScaleNoise = 0.00;
   P.clipLow = true;
   P.clipHigh = false;
   P.rangeClipLow = true;
   P.rangeLow = 0.000000;
   P.rangeClipHigh = false;
   P.rangeHigh = 0.980000;
   P.mapRangeRejection = true;
   P.reportRangeRejection = false;
   P.largeScaleClipLow = true;
   P.largeScaleClipLowProtectedLayers = 2;
   P.largeScaleClipLowGrowth = 3;
   P.largeScaleClipHigh = false;
   P.largeScaleClipHighProtectedLayers = 2;
   P.largeScaleClipHighGrowth = 2;
   P.generate64BitResult = false;
   P.generateRejectionMaps = true;
   P.generateSlopeMaps = false;
   P.generateIntegratedImage = true;
   P.generateDrizzleData = false;
   P.closePreviousImages = false;
   P.bufferSizeMB = 16;
   P.stackSizeMB = 1024;
   P.autoMemorySize = true;
   P.autoMemoryLimit = 0.75;
   P.useROI = false;
   P.roiX0 = 0;
   P.roiY0 = 0;
   P.roiX1 = 0;
   P.roiY1 = 0;
   P.useCache = true;
   P.evaluateSNR = false;
   P.noiseEvaluationAlgorithm = ImageIntegration.prototype.NoiseEvaluation_MRS;
   P.mrsMinDataFraction = 0.010;
   P.psfStructureLayers = 5;
   P.psfType = ImageIntegration.prototype.PSFType_Moffat4;
   P.generateFITSKeywords = true;
   P.subtractPedestals = false;
   P.truncateOnOutOfRange = false;
   P.noGUIMessages = true;
   P.showImages = true;
   P.useFileThreads = true;
   P.fileThreadOverload = 1.00;
   P.useBufferThreads = true;
   P.maxBufferThreads = 0;

   console.writeln( "<end><cbr><br>Integrating block ", block.blockNumber,
                    " (", block.inputFiles.length, " input frames) to:<br><raw>",
                    block.outputPath, "</raw>" );
   if ( !P.executeGlobal() )
      throw new Error( "ImageIntegration failed for block " + block.blockNumber +
                       ": " + block.outputPath );

   var integrationWindow = ImageWindow.windowById( P.integrationImageId );
   if ( integrationWindow.isNull )
      throw new Error( "ImageIntegration did not create an integration image." );
   if ( File.exists( block.outputPath ) )
      throw new Error( "Integration output appeared during processing and will not be overwritten: " + block.outputPath );

   integrationWindow.saveAs( block.outputPath, false, false, false, false );
   if ( !File.exists( block.outputPath ) )
      throw new Error( "Failed to save integration output: " + block.outputPath );

   integrationWindow.forceClose();
   closeWindowById( P.lowRejectionMapImageId );
   closeWindowById( P.highRejectionMapImageId );
   closeWindowById( P.slopeMapImageId );
   console.writeln( "<end><cbr>Saved integration output:<br><raw>", block.outputPath, "</raw>" );
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
         throw new Error( "Integration manifest not found: " + manifestPath );

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
         integrateBlock( block );
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
         console.criticalln( "<end><cbr>Unable to write integration status: ",
                             errorText( statusError ) );
      }
      console.criticalln( "<end><cbr>IntegrateTotality failed: ", status.errorMessage );
      throw error;
   }
}

main();
