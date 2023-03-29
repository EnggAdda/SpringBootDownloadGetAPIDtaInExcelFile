package com.example.springbootdownloadexcelfile.service.impl;


import com.example.springbootdownloadexcelfile.entity.Product;
import com.example.springbootdownloadexcelfile.repository.ProductRepo;
import com.example.springbootdownloadexcelfile.service.ProductService;
import com.example.springbootdownloadexcelfile.util.ExcelUtil;
import jakarta.annotation.PostConstruct;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Random;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Service
public class ProductServiceImpl implements ProductService {
   @Autowired
    private ProductRepo productRepo;

    @Override
    public List<Product> getAllProducts() {
        return productRepo.findAll();
    }

    @Override
    public Product insertProductIntoDatabase(Product product) {
        return productRepo.save(product);
    }
    @Override
    public Product getProductById(int id ) {
        return productRepo.findById(id).get();
    }
    @Override
    public List<Product> findProductsWithSorting(String field){
        return  productRepo.findAll(Sort.by(Sort.Direction.ASC,field));
    }
    @Override
    public Page<Product> findProductsWithPagination(int offset, int pageSize){
        Page<Product> products = productRepo.findAll(PageRequest.of(offset, pageSize));
        return  products;
    }
    @Override
    public Page<Product> findProductsWithPaginationAndSorting(int offset,int pageSize,String field){
        Page<Product> products = productRepo.findAll(PageRequest.of(offset, pageSize).withSort(Sort.by(field)));
        return  products;
    }

    @Override
    public List<Product> sortBasedOnSomeField(String field) {
        return productRepo.findAll(Sort.by(Sort.Direction.ASC,field));

    }

    @Override
    public Page<Product> getProductWithPagination(int offset, int pageSize) {
        return productRepo.findAll(PageRequest.of(offset, pageSize));
    }

    @Override
    public Page<Product> getProoductWithPaginationAndSorting(int offset, int pageSize, String field) {
        return productRepo.findAll(PageRequest.of(offset,pageSize).withSort(Sort.by(field)));
    }

    @Override
    public ByteArrayInputStream getDataDownload() throws IOException {
        List<Product> list = productRepo.findAll();
        ByteArrayInputStream byteArrayInputStream =
                ExcelUtil.dataToExcel(list);
        return byteArrayInputStream;
    }

    @PostConstruct
    public void initDB() {
        List<Product> products = IntStream.rangeClosed(1, 300)
                .mapToObj(i -> new Product(i,"product" + i, 1000.0, new Random().nextInt(50000)))
               .collect(Collectors.toList());
        productRepo.saveAll(products);
   }
}
