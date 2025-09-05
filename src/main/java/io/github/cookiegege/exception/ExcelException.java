package io.github.cookiegege.exception;

import cn.hutool.core.util.StrUtil;
import lombok.Getter;
import lombok.Setter;

/**
 * 通用异常
 *
 * @author xuyuxiang
 * @date 2020/4/8 15:54
 */
@Getter
@Setter
public class ExcelException extends RuntimeException {

    private Integer code;

    private String msg;

    public ExcelException() {
        super("Excel操作异常");
        this.code = 500;
        this.msg = "Excel操作异常";
    }

    public ExcelException(String msg, Object... arguments) {
        super(StrUtil.format(msg, arguments));
        this.code = 500;
        this.msg = StrUtil.format(msg, arguments);
    }

    public ExcelException(Integer code, String msg, Object... arguments) {
        super(StrUtil.format(msg, arguments));
        this.code = code;
        this.msg = StrUtil.format(msg, arguments);
    }
}
