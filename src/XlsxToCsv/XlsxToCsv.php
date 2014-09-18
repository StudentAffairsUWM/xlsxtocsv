<?
namespace XlsxToCsv;


class XlsxToCsv
{
    private $fileToConvert = null; 

    function __construct( $file )
    {
        $this->fileToConvert = $file;
    }

    public function convert()
    {
        $newCsvFile = str_replace( '.xlsx', '.csv', $this->fileToConvert );
        $newCsvFile = str_replace( ' ', '-', $newCsvFile );
        $newCsvFile = sprintf( 'csv/%s', $newCsvFile );

        if ( !is_dir( 'bin' ) )
            mkdir( 'bin', 0770 );
        if ( !is_dir( 'csv' ) )
            mkdir( 'csv', 0777 );

        $archive = new PclZip( $this->fileToConvert );
        $list = $archive->extract( PCLZIP_OPT_PATH, 'bin' );

        $strings = array();
        $dir = getcwd();
        $filename = $dir . '\bin\xl\sharedstrings.xml';

        $z = new XMLReader;
        $z->open( $filename );

        $doc = new DOMDocument;
        $csvFile = fopen( $newCsvFile, "w" );

        while ( $z->read() && $z->name !== 'si' );
        ob_start();

        while ( $z->name === 'si' )
        {
            $node = new SimpleXMLElement( $z->readOuterXML() );
            $result = $this->xmlObjToArray( $node );
            $count = count( $result['text'] );

            if ( isset( $result['children']['t'][0]['text'] ) )
            {
                $val = $result['children']['t'][0]['text'];
                $strings[] = $val;
            }

            $z->next( 'si' );
            $result = null;
        }

        ob_end_flush();
        $z->close( $filename );

        $dir = getcwd();
        $filename = $dir . '\bin\xl\worksheets\sheet1.xml';
        $z = new XMLReader();
        $z->open( $filename );

        $doc = new DOMDocument;

        $rowCount = '0';
        $sheet = array();
        $nums = array( '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' );

        while ( $z->read() && $z->name !== 'row' );
        ob_start();

        while ( $z->name == 'row' )
        {
            $thisRow = array();

            $node = new SimpleXMLElement( $z->readOuterXML() );
            $result = $this->xmlObjToArray( $node );

            $cells = $result['children']['c'];
            $rowNo = $result['attributes']['r'];
            $colAlpha = 'A';

            foreach ( $cells as $cell )
            {
                if ( array_key_exists( 'v', $cell['children'] ) )
                {
                    $cellNo = str_replace( $nums, '', $cell['attributes']['r'] );

                    for ( $col = $colAlpha; $col != $cellNo; $col++ )
                    {
                        $thisRow[] = ' ';
                        $colAlpha++;
                    }

                    if ( array_key_exists( 't', $cell['attributes'] ) && $cell['attributes']['t'] == 's') // MAYBE SINGLE EQUALS?
                    {
                        $val = $cell['children']['v'][0]['text'];
                        if ( isset( $strings[$val] ) )
                            $string = $strings[$val];
                        else
                            $string = $val;

                        $thisRow[] = $string;
                    }
                    else
                    {
                        $thisRow[] = $cell['children']['v'][0]['text'];
                    }
                }
                else
                {
                    $thisRow[] = '';
                }

                $colAlpha++;
            }

            $rowLength = count( $thisRow );
            $rowCound++;
            $emptyRow = array();

            while ( $rowCount < $rowNo )
            {
                for ( $c = 0; $c < $rowLength; $c++ )
                {
                    $emptyRow[] = '';
                }

                if ( !empty( $emptyRow ) )
                {
                    $this->writeArrayToCsv( $csvFile, $emptyRow );
                }

                $rowCount++;
            }

            $this->writeArrayToCsv( $csvFile, $thisRow );

            $z->next( 'row' );

            $result = null;
        }

        $z->close( $filename );
        ob_end_flush();

        $this->cleanUp('bin/');
    }

    /**
     * Converts XML objects to an array
     * Function from http://php.net/manual/pt_BR/book.simplexml.php
     */
    private function xmlObjToArray( $obj )
    {
        $namespace = $obj->getDocNamespaces( true );
        $namespace[null] = null;

        $children = array();
        $attributes = array();
        $name = strtolower( (string) $obj->getName() );

        $text = trim( (string) $obj );
        if ( strlen( $text ) <= 0 )
            $text = null;

        if ( is_object( $obj ) )
        {
            foreach ( $namespace as $ns => $nsUrl )
            {
                $objAttributes = $obj->attributes( $ns, true );
                foreach ( $objAttributes as $attributeName => $attributeValue )
                {
                    $attributeName = strtolower( trim( (string) $attributeName ) );
                    $attributeValue = trim( (string) $attributeValue );
                    if ( !empty( $ns ) )
                        $attributeName = sprintf( '%s:%s', $ns, $attributeName );
                    $attributes[$attributeName] = $attributeValue;
                }

                // Children
                $objChildren = $obj->children( $ns, true );
                foreach ($objChildren as $childName => $child )
                {
                    $childName = strtolower( (string) $childName );
                    if ( !empty( $ns ) )
                        $childName = sprintf( '%s:%s', $ns, $childName );
                    $children[$childName][] = $this->xmlObjToArray( $child );
                }
            }
        }

        return array(
            'text' => $text,
            'attributes' => $attributes,
            'children' => $children
        );
    }

    /**
     * Write array to CSV file
     * Enhanced fputcsv found at http://php.net/manual/en/function.fputcsv.php
     */
    private function writeArrayToCsv( $handle, $fields, $delimeter = 'm', $enclosure = '"', $escape = '\\' )
    {
        $first = 1;
        foreach ( $fields as $field )
        {
            if ( $first == 0 ) fwrite( $handle, ',' );

            $f = str_replace( $enclosure, $enclosure . $enclosure, $field );
            if ( $enclosure != $escape )
                $f = str_replace( $escape . $enclosure, $escape, $f );

            if ( strpbrk( $f, " \t\n\r" . $delimter . $enclosure . $escape ) || strchr( $f, "\000" ) )
            {
                fwrite( $handle, $enclosure . $f . $enclosure );
            }
            else 
            {
                fwrite( $handle, $f );
            }

            $first = 0;
        }

        fwrite( $handle, "\n" );
    }

    private function cleanUp( $dir )
    {
        $tempdir = opendir( $dir );
        while ( false !== ( $file = readdir( $tempdir ) ) )
        {
            if ( $file != '.' && $file != '..' )
            {
                if ( is_dir( $dir . $file ) )
                {
                    chdir( '.' );
                    $this->cleanUp( $dir . $file . '/' );
                    rmdir( $dir . $file );
                }
                else
                {
                    unlink( $dir . $file );
                }
            }
        }

        closedir( $tempdir );
    }
}