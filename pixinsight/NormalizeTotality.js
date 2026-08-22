// ----------------------------------------------------------------------------
// NormalizeTotality.js
// Normalizes Totality HDR frames by a reference preview and applies fixed RGB
// color coefficients derived from the same preview.
// ----------------------------------------------------------------------------

#feature-id    Utilities > NormalizeTotality
#feature-info  Normalizes and color-corrects aligned Totality HDR blocks prepared by AsiToPix.

#include <pjsr/ImageOp.jsh>
#include <pjsr/UndoFlag.jsh>

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

function canonicalPath( path )
{
   return File.fullPath( File.windowsPathToUnix( path ) ).toLowerCase();
}

function validateManifest( manifest, manifestPath )
{
   if ( manifest === null || typeof manifest != "object" )
      throw new Error( "Invalid Totality normalization manifest: " + manifestPath );
   if ( manifest.schemaVersion != 1 )
      throw new Error( "Unsupported Totality normalization manifest schema: " + manifestPath );
   if ( typeof manifest.statusPath != "string" || manifest.statusPath.length == 0 )
      throw new Error( "The Totality normalization manifest has no status path: " + manifestPath );
   if ( typeof manifest.metricsPath != "string" || manifest.metricsPath.length == 0 )
      throw new Error( "The Totality normalization manifest has no metrics path: " + manifestPath );
   if ( !(manifest.allowedReferenceFiles instanceof Array) || manifest.allowedReferenceFiles.length == 0 )
      throw new Error( "The Totality normalization manifest has no allowed reference files: " + manifestPath );
   if ( !(manifest.frames instanceof Array) || manifest.frames.length == 0 )
      throw new Error( "The Totality normalization manifest has no frames: " + manifestPath );
   if ( File.directoryExists( manifest.metricsPath ) )
      throw new Error( "Normalization metrics path is a directory: " + manifest.metricsPath );

   for ( var i = 0; i < manifest.frames.length; ++i )
   {
      var frame = manifest.frames[i];
      if ( frame === null || typeof frame != "object" || typeof frame.blockNumber != "number" )
         throw new Error( "Invalid Totality normalization frame in manifest: " + manifestPath );
      if ( typeof frame.inputPath != "string" || !File.exists( frame.inputPath ) )
         throw new Error( "Totality normalization input frame not found: " + frame.inputPath );
      if ( typeof frame.normalizedPath != "string" || frame.normalizedPath.length == 0 ||
           typeof frame.colorCorrectedPath != "string" || frame.colorCorrectedPath.length == 0 )
         throw new Error( "Totality normalization block " + frame.blockNumber + " has invalid output paths." );
      if ( File.directoryExists( frame.normalizedPath ) )
         throw new Error( "Normalization output path is a directory: " + frame.normalizedPath );
      if ( File.directoryExists( frame.colorCorrectedPath ) )
         throw new Error( "Color-corrected output path is a directory: " + frame.colorCorrectedPath );
   }
}

function selectedReferencePreview()
{
   var window = ImageWindow.activeWindow;
   if ( window.isNull )
      throw new Error( "No active reference image. Open a composed Totality frame and create a preview first." );
   if ( typeof window.filePath != "string" || window.filePath.length == 0 )
      throw new Error( "The active reference image has no file path. Open a saved Block-*-HDR.xisf frame." );

   var preview = window.currentView;
   if ( !preview.isPreview )
   {
      var previews = window.previews;
      if ( previews.length != 1 )
         throw new Error( "Activate the intended reference preview, or leave exactly one preview in the reference image." );
      preview = previews[0];
   }
   return {
      window: window,
      view: preview,
      rect: window.previewRect( preview )
   };
}

function validateReferencePath( path, allowedReferenceFiles )
{
   var referencePath = canonicalPath( path );
   for ( var i = 0; i < allowedReferenceFiles.length; ++i )
      if ( referencePath == canonicalPath( allowedReferenceFiles[i] ) )
         return;
   throw new Error( "The active reference image is not one of the composed input frames: " + path );
}

function measureRgbMedians( image, rect, description )
{
   if ( image.numberOfChannels < 3 || !image.isColor )
      throw new Error( "Expected an RGB image for " + description + "." );
   if ( rect.x0 < 0 || rect.y0 < 0 || rect.x1 > image.width || rect.y1 > image.height ||
        rect.width <= 0 || rect.height <= 0 )
      throw new Error( "Reference preview ROI is outside " + description + "." );

   var medians = [];
   try
   {
      image.selectedRect = rect;
      for ( var channel = 0; channel < 3; ++channel )
      {
         image.selectedChannel = channel;
         var median = image.median();
         if ( !isFinite( median ) || median <= 0 )
            throw new Error( "Invalid or zero channel median in " + description +
                             " for channel " + channel + ": " + median );
         medians.push( median );
      }
   }
   finally
   {
      image.resetSelections();
   }
   return medians;
}

function median3( a, b, c )
{
   var values = [ a, b, c ];
   values.sort( function( x, y ) { return x - y; } );
   return values[1];
}

function multiplyImage( view, scalar )
{
   view.beginProcess( UndoFlag_NoSwapFile );
   try
   {
      view.image.resetSelections();
      view.image.apply( scalar, ImageOp_Mul );
   }
   finally
   {
      view.endProcess();
   }
}

function multiplyChannels( view, coefficients )
{
   view.beginProcess( UndoFlag_NoSwapFile );
   try
   {
      view.image.resetSelections();
      for ( var channel = 0; channel < 3; ++channel )
      {
         view.image.selectedChannel = channel;
         view.image.apply( coefficients[channel], ImageOp_Mul );
      }
      view.image.resetSelections();
   }
   finally
   {
      view.endProcess();
   }
}

function openSingleImage( path )
{
   var windows = ImageWindow.open( path );
   if ( windows.length != 1 )
   {
      for ( var i = 0; i < windows.length; ++i )
         if ( !windows[i].isNull )
            windows[i].forceClose();
      throw new Error( "Expected one image in input file, found " + windows.length + ": " + path );
   }
   return windows[0];
}

function csvField( value )
{
   var text = String( value );
   return '"' + text.replace( /"/g, '""' ) + '"';
}

function numberText( value )
{
   return Number( value ).toPrecision( 17 );
}

function metricsHeader()
{
   return [
      "Block", "InputFile", "Reference", "PreviewId",
      "RoiX0", "RoiY0", "RoiX1", "RoiY1",
      "Rref", "Gref", "Bref", "R", "G", "B",
      "kR", "kG", "kB", "k", "cR", "cG", "cB",
      "NormalizedFile", "ColorCorrectedFile"
   ].join( "," );
}

function metricsRow( frame, isReference, previewId, rect, referenceMedians,
                     medians, factors, scalar, colorCoefficients )
{
   return [
      frame.blockNumber,
      csvField( frame.inputPath ),
      isReference ? "true" : "false",
      csvField( previewId ),
      rect.x0, rect.y0, rect.x1, rect.y1,
      numberText( referenceMedians[0] ),
      numberText( referenceMedians[1] ),
      numberText( referenceMedians[2] ),
      numberText( medians[0] ),
      numberText( medians[1] ),
      numberText( medians[2] ),
      numberText( factors[0] ),
      numberText( factors[1] ),
      numberText( factors[2] ),
      numberText( scalar ),
      numberText( colorCoefficients[0] ),
      numberText( colorCoefficients[1] ),
      numberText( colorCoefficients[2] ),
      csvField( frame.normalizedPath ),
      csvField( frame.colorCorrectedPath )
   ].join( "," );
}

function processFrame( frame, referencePath, previewId, rect,
                       referenceMedians, colorCoefficients )
{
   var window = null;
   try
   {
      window = openSingleImage( frame.inputPath );
      var view = window.mainView;
      var medians = measureRgbMedians( view.image, rect, frame.inputPath );
      var factors = [
         referenceMedians[0] / medians[0],
         referenceMedians[1] / medians[1],
         referenceMedians[2] / medians[2]
      ];
      var scalar = median3( factors[0], factors[1], factors[2] );
      if ( !isFinite( scalar ) || scalar <= 0 )
         throw new Error( "Invalid scalar normalization factor for block " + frame.blockNumber + ": " + scalar );

      multiplyImage( view, scalar );
      window.saveAs( frame.normalizedPath, false, false, false, false );
      if ( !File.exists( frame.normalizedPath ) )
         throw new Error( "Failed to save normalized frame: " + frame.normalizedPath );

      multiplyChannels( view, colorCoefficients );
      window.saveAs( frame.colorCorrectedPath, false, false, false, false );
      if ( !File.exists( frame.colorCorrectedPath ) )
         throw new Error( "Failed to save normalized color-corrected frame: " + frame.colorCorrectedPath );

      console.writeln( "<end><cbr>Block ", frame.blockNumber,
                       ": R=", numberText( medians[0] ),
                       ", G=", numberText( medians[1] ),
                       ", B=", numberText( medians[2] ),
                       ", k=", numberText( scalar ) );
      return metricsRow(
         frame,
         canonicalPath( frame.inputPath ) == referencePath,
         previewId,
         rect,
         referenceMedians,
         medians,
         factors,
         scalar,
         colorCoefficients
      );
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
      currentBlockNumber: null,
      completedBlockNumbers: [],
      referencePath: "",
      previewId: "",
      roi: null,
      referenceMedians: null,
      errorMessage: "",
      errorStack: ""
   };

   try
   {
      if ( !File.exists( manifestPath ) )
         throw new Error( "Totality normalization manifest not found: " + manifestPath );
      var manifest = JSON.parse( File.readTextFile( manifestPath ) );
      if ( manifest !== null && typeof manifest == "object" &&
           typeof manifest.statusPath == "string" )
         statusPath = manifest.statusPath;
      writeStatus( statusPath, status );
      validateManifest( manifest, manifestPath );

      var reference = selectedReferencePreview();
      validateReferencePath( reference.window.filePath, manifest.allowedReferenceFiles );
      var referencePath = canonicalPath( reference.window.filePath );
      var referenceMedians = measureRgbMedians(
         reference.window.mainView.image,
         reference.rect,
         "reference frame " + reference.window.filePath
      );
      var colorCoefficients = [
         referenceMedians[1] / referenceMedians[0],
         1,
         referenceMedians[1] / referenceMedians[2]
      ];

      status.state = "running";
      status.referencePath = reference.window.filePath;
      status.previewId = reference.view.id;
      status.roi = {
         x0: reference.rect.x0,
         y0: reference.rect.y0,
         x1: reference.rect.x1,
         y1: reference.rect.y1
      };
      status.referenceMedians = referenceMedians;
      writeStatus( statusPath, status );

      console.writeln( "<end><cbr><br>Reference: <raw>", reference.window.filePath, "</raw>" );
      console.writeln( "Preview ", reference.view.id, ": ",
                       reference.rect.x0, ",", reference.rect.y0, " - ",
                       reference.rect.x1, ",", reference.rect.y1 );
      console.writeln( "Rref=", numberText( referenceMedians[0] ),
                       ", Gref=", numberText( referenceMedians[1] ),
                       ", Bref=", numberText( referenceMedians[2] ) );
      console.writeln( "cR=", numberText( colorCoefficients[0] ),
                       ", cG=1, cB=", numberText( colorCoefficients[2] ) );

      var csvLines = [ metricsHeader() ];
      for ( var i = 0; i < manifest.frames.length; ++i )
      {
         var frame = manifest.frames[i];
         status.currentBlockNumber = frame.blockNumber;
         writeStatus( statusPath, status );
         csvLines.push( processFrame(
            frame,
            referencePath,
            reference.view.id,
            reference.rect,
            referenceMedians,
            colorCoefficients
         ) );
         status.completedBlockNumbers.push( frame.blockNumber );
         writeStatus( statusPath, status );
      }

      File.writeTextFile( manifest.metricsPath, csvLines.join( "\n" ) + "\n" );
      if ( !File.exists( manifest.metricsPath ) )
         throw new Error( "Failed to save normalization metrics CSV: " + manifest.metricsPath );

      status.state = "completed";
      status.currentBlockNumber = null;
      writeStatus( statusPath, status );
      console.writeln( "<end><cbr><br>Saved normalization metrics:<br><raw>",
                       manifest.metricsPath, "</raw>" );
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
         console.criticalln( "<end><cbr>Unable to write normalization status: ",
                             errorText( statusError ) );
      }
      console.criticalln( "<end><cbr>NormalizeTotality failed: ", status.errorMessage );
      throw error;
   }
}

main();
