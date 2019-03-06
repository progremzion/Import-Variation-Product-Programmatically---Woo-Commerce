<?php
//Set Maximum Execution Time - To Avoid Execution
ini_set('max_execution_time', 0);

//Load WordPress
require_once("wp-load.php"); // Include the class_alias()
require_once(ABSPATH . 'wp-admin/includes/taxonomy.php');

include 'excel_reader/excel_reader.php'; // Include the class
$excel = new PhpExcelReader; // Creates object instance of the class
$excel->read('import/import.xls'); // Reads and stores the excel file data

//Define Variable
$categoryIDS = array();
$masterArray = array();

//Create category in system
$cats = array(
    array('thumb' => '','name' => 'Confort','description' => '','slug' => '','parent' => ''),
    array('thumb' => '','name' => 'Caballero','description' => '','slug' => '','parent' => ''),
    array('thumb' => '','name' => 'Sixpack','description' => '','slug' => '','parent' => '')
);

//Process category array
foreach($cats as $data) {

	$term = term_exists( $data['name'], 'product_cat'); //Check category exist

	if ( isset($term['term_id']) && $term['term_id'] != '' ) {
		$categoryIDS[$data['name']] = $term['term_id'];
	} else {

		$cid = wp_insert_term(
	        $data['name'], // the term
	        'product_cat', // the taxonomy
	        array(
	            'description'=> $data['description'],
	            'slug' => $data['slug'],
	            'parent' => $data['parent']
	        )
	    );

	    $cat_id = isset( $cid['term_id'] ) ? $cid['term_id'] : 0;

	    $categoryIDS[$data['name']] = $cat_id; //Update category id

	    if (isset($data['thumb'])) {
			update_woocommerce_term_meta( $cid['term_id'], 'thumbnail_id', absint( $thumb_id ) );
		}
	}
}


//Master array
foreach ($excel->sheets[0]['cells'] as $key => $value) {
	if ($value[1] != 'MODEL #'):

	$masterArray[$value[1]]['color'][$value['2']][] = $value['4'];
	$masterArray[$value[1]]['description'][] = $value['5'];
	$masterArray[$value[1]]['subcategory'][] = $value['6'];
	$masterArray[$value[1]]['discount'][] = $value['8'];
	$masterArray[$value[1]]['price'][] = $value['7'];
	$masterArray[$value[1]]['weight'][] = $value['9'];
	$masterArray[$value[1]]['shortdes'][] = $value['11'];
	$masterArray[$value[1]]['sku'][$value['2']][] = $value['12'];
	$masterArray[$value[1]]['category'][] = $categoryIDS[$value['13']];

	endif;
}


//Insert variant product
foreach ($masterArray as $key => $value) {

	echo '== MODEL =='.'<br>';
	echo $key.'<br>';

	$objProduct = new WC_Product_Variable();

	$objProduct->set_name($key.' '.htmlentities($value['description'][0]));
	$objProduct->set_status("publish");  // can be publish,draft or any wordpress post status
	$objProduct->set_catalog_visibility('visible'); // add the product visibility status
	$objProduct->set_description(htmlentities($value['description'][0]));
	$objProduct->set_sku(''); //can be blank in case you don't have sku, but You can't add duplicate sku's
	$objProduct->set_price($value['price'][0]); // set product price
	$objProduct->set_regular_price($value['price'][0]); // set product regular price
	$objProduct->set_manage_stock(false); // true or false
	$objProduct->set_stock_quantity('10');
	$objProduct->set_stock_status('instock'); // in stock or out of stock value
	$objProduct->set_backorders('no');
	$objProduct->set_reviews_allowed(true);
	$objProduct->set_short_description(htmlentities($value['shortdes'][0]));
	$objProduct->set_sold_individually(false);
	$objProduct->set_category_ids(array($value['category'][0])); // array of category ids, You can get category id from WooCommerce Product Category Section of Wordpress Admin

	$product_id = $objProduct->save(); // it will save the product and return the generated product id


	//Set the Tag
	wp_set_post_terms( $product_id, $value['subcategory'][0], 'product_tag', true );

	echo '== Product Tag Added =='.'<br>';

	$img = str_replace('-', '', $key);
	$image = 'img/'.$img.'_C.png'; // images url array of product

	uploadMedia($image,$product_id); // calling the uploadMedia function and passing image url to get the uploaded media id

	echo '== Attribute Process Begin =='.'<br>';
	foreach ($value['color'] as $k => $v) {

		$term_taxonomy_ids = wp_set_object_terms( $product_id, $k, 'pa_color', true );
		$thedata = Array(
		     'pa_color'=>Array(
		           'name'=>'pa_color',
		           'value'=> $k,
		           'is_visible' => '1',
		           'is_variation' => '1',
		           'is_taxonomy' => '1'
		     )
		);

		//First getting the Post Meta
    	$_product_attributes = get_post_meta($product_id, '_product_attributes', TRUE);

		update_post_meta( $product_id,'_product_attributes',array_merge($_product_attributes, $thedata));

		foreach ($v as $kv => $vv) {
			$term_taxonomy_ids = wp_set_object_terms( $product_id, $vv, 'pa_size', true );
			$thedata = Array(
			     'pa_size'=>Array(
			           'name'=>'pa_size',
			           'value'=> $vv,
			           'is_visible' => '1',
			           'is_variation' => '1',
			           'is_taxonomy' => '1'
			     )
			);

			//First getting the Post Meta
	    	$_product_attributes = get_post_meta($product_id, '_product_attributes', TRUE);

			update_post_meta( $product_id,'_product_attributes',array_merge($_product_attributes, $thedata));
		}
	}

	foreach ($value['color'] as $ck => $cv) {
		$variationArray = array();
		foreach ($cv as $k => $v) {

			//Prepare an array and pass it to function
			$variationArray = array(
			    'attributes' => array(
			        'size'  => $v,
			        'color' => $ck,
			    ),
			    'sku'           => $value['sku'][$ck][$k],
			    'regular_price' => $value['price'][$k],
			    'sale_price'    => '',
			    'stock_qty'     => 10,
			);

			// The function to be run
			create_product_variation( $product_id, $variationArray );
		}
	}
}

// Upload image to wordpress and returns the media id
function uploadMedia($image_url,$post_id){
	require_once(ABSPATH . 'wp-admin/includes/image.php');
	require_once(ABSPATH . 'wp-admin/includes/file.php');
	require_once(ABSPATH . 'wp-admin/includes/media.php');

	$upload_dir = wp_upload_dir(null);
    $image_data = file_get_contents($image_url);
    $attach_id = '';

    if ($image_data) {
    	$filename = basename($image_url);
	    if(wp_mkdir_p($upload_dir['path']))     $file = $upload_dir['path'] . '/' . $filename;
	    else                                    $file = $upload_dir['basedir'] . '/' . $filename;
	    file_put_contents($file, $image_data);
	    $wp_filetype = wp_check_filetype($filename, null );
	    $attachment = array(
	        'post_mime_type' => $wp_filetype['type'],
	        'post_title' => sanitize_file_name($filename),
	        'post_content' => '',
	        'post_status' => 'inherit'
	    );
	    $attach_id = wp_insert_attachment( $attachment, $file, $post_id );

	    $attach_data = wp_generate_attachment_metadata( $attach_id, $file );
	    $res1= wp_update_attachment_metadata( $attach_id, $attach_data );
	    $res2= set_post_thumbnail( $post_id, $attach_id );
    }

    return $attach_id;

}


function create_product_variation( $product_id, $variation_data ){

	require_once(ABSPATH . 'wp-admin/includes/taxonomy.php');

    // Get the Variable product object (parent)
    $product = wc_get_product($product_id);

    $variation_post = array(
        'post_title'  => $product->get_title(),
        'post_name'   => 'product-'.$product_id.'-variation',
        'post_status' => 'publish',
        'post_parent' => $product_id,
        'post_type'   => 'product_variation',
        'guid'        => $product->get_permalink()
    );

    // Creating the product variation
    $variation_id = wp_insert_post( $variation_post );

    // Get an instance of the WC_Product_Variation object
    $variation = new WC_Product_Variation( $variation_id );

    // Iterating through the variations attributes
    foreach ($variation_data['attributes'] as $attribute => $term_name )
    {
        $taxonomy = 'pa_'.$attribute; // The attribute taxonomy

        // If taxonomy doesn't exists we create it (Thanks to Carl F. Corneil)
        if( ! taxonomy_exists( $taxonomy ) ){

            register_taxonomy(
                $taxonomy,
               'product_variation',
                array(
                    'hierarchical' => false,
                    'label' => ucfirst( $taxonomy ),
                    'query_var' => true,
                    'rewrite' => array( 'slug' => '$taxonomy'), // The base slug
                )
            );
        }

        // Check if the Term name exist and if not we create it.
        if( ! term_exists( $term_name, $taxonomy ) )
            wp_insert_term( $term_name, $taxonomy ); // Create the term

        $term_slug = get_term_by('name', $term_name, $taxonomy )->slug; // Get the term slug

        // Get the post Terms names from the parent variable product.
        $post_term_names =  wp_get_post_terms( $product_id, $taxonomy, array('fields' => 'names') );

        // Check if the post term exist and if not we set it in the parent variable product.
        if( ! in_array( $term_name, $post_term_names ) )
            wp_set_post_terms( $product_id, $term_name, $taxonomy, true );

        // Set/save the attribute data in the product variation
        update_post_meta( $variation_id, 'attribute_'.$taxonomy, $term_slug );
    }

    ## Set/save all other data

    // SKU
    if( ! empty( $variation_data['sku'] ) )
        $variation->set_sku( $variation_data['sku'] );

    // Prices
    if( empty( $variation_data['sale_price'] ) ){
        $variation->set_price( $variation_data['regular_price'] );
    } else {
        $variation->set_price( $variation_data['sale_price'] );
        $variation->set_sale_price( $variation_data['sale_price'] );
    }
    $variation->set_regular_price( $variation_data['regular_price'] );

    // Stock
    if( ! empty($variation_data['stock_qty']) ){
        $variation->set_stock_quantity( $variation_data['stock_qty'] );
        $variation->set_manage_stock(true);
        $variation->set_stock_status('');
    } else {
        $variation->set_manage_stock(false);
    }

    $variation->set_weight(''); // weight (reseting)

    $variation->save(); // Save the data
}