package com.qa.opencart.tests;

import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.qa.opencart.base.BaseTest;

public class CheckOutPageTest extends BaseTest {

	@BeforeClass
	public void checkoutInfoSetup() {
		accPage = loginPage.doLogin(prop.getProperty("username"), prop.getProperty("pwd"));
	}

	@Test
	public void productHeaderTest() throws InterruptedException {
		resultPage = accPage.doSearch("macbook");
		productInfoPage = resultPage.selectProduct("MacBook Pro");
		productInfoPage.selectAddCart();
		//Thread.sleep(10000);
		productInfoPage.selectViewItems();
		checkoutPage = productInfoPage.selectViewCarts();

		Assert.assertEquals(checkoutPage.getShoppingCardHeaderName(), "Shopping Cart Checkout");
	}

}
