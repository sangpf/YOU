package com.youchu.product.service.impl;

import com.youchu.product.dao.ProductDao;
import com.youchu.product.entity.Product;
import com.youchu.product.service.ProductService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class ProductServiceImpl implements ProductService {

	@Autowired
	private ProductDao productDao;

	@Override
	public Product findById(Long id) {
		return productDao.findById(id).get();
	}

	@Override
	public void save(Product product) {
		productDao.save(product);
	}

	@Override
	public void update(Product product) {
		productDao.save(product);
	}

	@Override
	public void delete(Long id) {
		productDao.deleteById(id);
	}
}
