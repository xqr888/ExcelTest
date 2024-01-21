package entity;

import java.io.Serializable;
import java.util.Date;
public class Product implements Serializable {
    Integer id;
    String name;
    Double price;
    Integer count;
    Date createTime;

    public Product() {
    }

    public Product(Integer id, String name, Double price, Integer count, Date createTime) {
        this.id = id;
        this.name = name;
        this.price = price;
        this.count = count;
        this.createTime = createTime;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Double getPrice() {
        return price;
    }

    public void setPrice(Double price) {
        this.price = price;
    }

    public Integer getCount() {
        return count;
    }

    public void setCount(Integer count) {
        this.count = count;
    }

    public Date getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }

    @Override
    public String toString() {
        return "Product{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", price=" + price +
                ", count=" + count +
                ", createTime=" + createTime +
                '}';
    }
}
