package com.example.springbootdownloadexcelfile.service;



import com.example.springbootdownloadexcelfile.entity.Product;
import org.springframework.data.domain.Page;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;


public interface ProductService {

    List<Product> getAllProducts();

    Product insertProductIntoDatabase(Product product);

    Product getProductById(int id );

    List<Product> findProductsWithSorting(String field);

    Page<Product> findProductsWithPagination(int offset, int pageSize);

    Page<Product> findProductsWithPaginationAndSorting(int offset,int pageSize,String field);

    List<Product> sortBasedOnSomeField( String field);

    Page<Product> getProductWithPagination(int offset , int pageSize);

    Page<Product>  getProoductWithPaginationAndSorting(int offset, int pageSize, String field);


}
