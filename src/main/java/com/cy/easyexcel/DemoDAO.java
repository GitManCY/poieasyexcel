package com.cy.easyexcel;

import java.util.List;

/**
 * @description:
 * @projectName:poieasyexcel
 * @see:com.cy.easyexcel
 * @author:chengyang
 * @createTime:2020/4/29 11:28 上午
 * @version:1.0
 */
public class DemoDAO {
    /**
     * 假设这个是你的DAO存储。当然还要这个类让spring管理，当然你不用需要存储，也不需要这个类。
     **/
    public void save(List<DemoData> list) {
        // 如果是mybatis,尽量别直接调用多次insert,自己写一个mapper里面新增一个方法batchInsert,所有数据一次性插入
    }

}
