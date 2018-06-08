module.exports = {
  /**
   * sheet(table) 的类型
   * 影响输出json的类型
   * 当有#id类型的时候  表输出json的是map形式(id:{xx:1})
   * 当没有#id类型的时候  表输出json的是数组类型 没有id索引
   */
  SheetType: {
    /**
     * 普通表 
     * 输出JSON ARRAY
     */
    NORMAL: 0,

    /**
     * 有主外键关系的主表
     * 输出JSON MAP
     */
    MASTER: 1,

    /**
     * 有主外键关系的附表
     * 输出JSON MAP
     */
    SLAVE: 2
  },

  /**
   * 支持的数据类型
   */
  DataType: {
    NUMBER: 'number',
    STRING: 'string',
    BOOL: 'bool',
    DATE: 'date',
    ID: 'id',
    IDS: 'id[]',
    ARRAY: '[]',
    OBJECT: '{}',
    UNKNOWN: 'unknown'
  }

};