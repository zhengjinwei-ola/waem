# 多阶段构建Dockerfile
FROM public.ecr.aws/docker/library/rust:1.75-slim as builder

WORKDIR /app

# 配置 Cargo 使用中科大镜像（通过环境变量，更稳定）
ENV CARGO_HOME=/usr/local/cargo
ENV CARGO_REGISTRIES_CRATES_IO_INDEX=https://mirrors.ustc.edu.cn/crates.io-index
ENV CARGO_REGISTRIES_CRATES_IO_PROTOCOL=sparse

# 仅复制Cargo.toml（避免无Cargo.lock时报错）
COPY Cargo.toml ./

# 创建空的src目录并为所有bin创建占位main以预构建依赖（缓存）
RUN mkdir -p src \
 && echo "fn main() {}" > src/main.rs \
 && echo "fn main() {}" > src/server.rs \
 && echo "fn main() {}" > src/generate_sample.rs \
 && cargo build --release \
 && rm -rf src

# 复制源代码
COPY . .

# 确保存在config目录（即使为空）以便后续复制
RUN mkdir -p config

# 安装构建依赖
RUN apt-get update && apt-get install -y \
    pkg-config \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# 构建应用
RUN cargo build --release --bin server

# 运行时镜像
FROM public.ecr.aws/debian/debian:bookworm-slim

# 配置 Debian 源为中科大镜像（使用更稳定的方式）
RUN echo "deb https://mirrors.ustc.edu.cn/debian bookworm main" > /etc/apt/sources.list && \
    echo "deb https://mirrors.ustc.edu.cn/debian-security bookworm-security main" >> /etc/apt/sources.list && \
    echo "deb https://mirrors.ustc.edu.cn/debian bookworm-updates main" >> /etc/apt/sources.list

# 安装运行时依赖
RUN apt-get update && apt-get install -y \
    ca-certificates \
    curl \
    && rm -rf /var/lib/apt/lists/*

# 创建应用用户
RUN useradd -r -s /bin/false -d /app app

# 设置工作目录
WORKDIR /app

# 复制二进制文件
COPY --from=builder /app/target/release/server /app/
COPY --from=builder /app/target/release/excel_to_word /app/
COPY --from=builder /app/target/release/generate_sample /app/

# 复制配置文件
COPY --from=builder /app/config /app/config/

# 设置权限
RUN chown -R app:app /app && chmod +x /app/server

# 切换到应用用户
USER app

# 暴露端口
EXPOSE 3002

# 设置环境变量
ENV RUST_LOG=info
ENV PORT=3002

# 健康检查
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:3002/ || exit 1

# 启动命令
CMD ["./server"]
