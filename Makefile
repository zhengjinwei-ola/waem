.PHONY: build run clean sample test help

# 默认目标
all: build

# 构建项目
build:
	@echo "正在构建项目..."
	cargo build --release
	@echo "构建完成！"

# 运行程序（需要提供输入文件）
run:
	@echo "使用方法: make run INPUT=your_file.csv [OUTPUT=output.docx]"
	@if [ -z "$(INPUT)" ]; then \
		echo "错误: 请指定输入文件，例如: make run INPUT=sample_bills.csv"; \
		exit 1; \
	fi
	@if [ -z "$(OUTPUT)" ]; then \
		cargo run --release -- --input $(INPUT); \
	else \
		cargo run --release -- --input $(INPUT) --output $(OUTPUT); \
	fi

# 生成示例数据
sample:
	@echo "正在生成示例CSV文件..."
	cargo run --release --bin generate_sample
	@echo "示例文件生成完成！"

# 测试完整流程
test: sample
	@echo "正在测试CSV到Word转换..."
	cargo run --release -- --input sample_bills.csv --output test_output.docx
	@echo "测试完成！请检查 test_output.docx 文件"

# 清理生成的文件
clean:
	@echo "正在清理生成的文件..."
	rm -f sample_bills.csv
	rm -f *.docx
	rm -f test_output.docx
	cargo clean
	@echo "清理完成！"

# 安装依赖
install:
	@echo "正在安装依赖..."
	cargo build
	@echo "依赖安装完成！"

# 显示帮助信息
help:
	@echo "可用的命令:"
	@echo "  build   - 构建项目"
	@echo "  run     - 运行程序 (需要指定 INPUT 参数)"
	@echo "  sample  - 生成示例CSV文件"
	@echo "  test    - 测试完整流程"
	@echo "  clean   - 清理生成的文件"
	@echo "  install - 安装依赖"
	@echo "  help    - 显示此帮助信息"
	@echo ""
	@echo "示例用法:"
	@echo "  make sample                    # 生成示例数据"
	@echo "  make run INPUT=sample_bills.csv    # 转换示例文件"
	@echo "  make test                     # 完整测试流程"
