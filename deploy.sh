#!/bin/bash

# Excel到Word转换器部署脚本
# 适用于Linux服务器部署

set -e  # 遇到错误立即退出

# 配置变量
APP_NAME="excel_to_word_server"
APP_VERSION="1.0.0"
BUILD_DIR="build"
SERVICE_NAME="excel-to-word"
USER_NAME="excel-word"
PORT=3002

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 日志函数
log_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

log_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# 检查是否为root用户
check_root() {
    if [[ $EUID -eq 0 ]]; then
        log_error "请不要使用root用户运行此脚本"
        exit 1
    fi
}

# 检查系统依赖
check_dependencies() {
    log_info "检查系统依赖..."
    
    # 检查Rust
    if ! command -v cargo &> /dev/null; then
        log_error "未找到Rust，正在安装..."
        curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
        source ~/.cargo/env
    fi
    
    # 检查其他必要工具
    local missing_tools=()
    for tool in systemctl journalctl; do
        if ! command -v $tool &> /dev/null; then
            missing_tools+=($tool)
        fi
    done
    
    if [[ ${#missing_tools[@]} -gt 0 ]]; then
        log_error "缺少必要工具: ${missing_tools[*]}"
        log_error "请确保在支持systemd的Linux发行版上运行"
        exit 1
    fi
    
    log_success "系统依赖检查完成"
}

# 构建应用
build_app() {
    log_info "开始构建应用..."
    
    # 清理旧的构建目录
    if [[ -d "$BUILD_DIR" ]]; then
        rm -rf "$BUILD_DIR"
    fi
    
    mkdir -p "$BUILD_DIR"
    
    # 构建release版本
    log_info "编译Rust项目..."
    cargo build --release
    
    if [[ $? -ne 0 ]]; then
        log_error "构建失败"
        exit 1
    fi
    
    # 复制二进制文件
    cp target/release/server "$BUILD_DIR/${APP_NAME}"
    cp target/release/excel_to_word "$BUILD_DIR/"
    cp target/release/generate_sample "$BUILD_DIR/"
    
    # 复制配置文件
    if [[ -d "config" ]]; then
        cp -r config "$BUILD_DIR/"
    fi
    
    # 创建部署信息
    cat > "$BUILD_DIR/deploy_info.txt" << EOF
部署时间: $(date)
应用版本: ${APP_VERSION}
构建主机: $(hostname)
Rust版本: $(rustc --version)
EOF
    
    log_success "应用构建完成"
}

# 创建系统用户
create_user() {
    log_info "创建系统用户..."
    
    if ! id "$USER_NAME" &>/dev/null; then
        sudo useradd -r -s /bin/false -d /opt/$APP_NAME $USER_NAME
        log_success "用户 $USER_NAME 创建成功"
    else
        log_warning "用户 $USER_NAME 已存在"
    fi
}

# 安装应用
install_app() {
    log_info "安装应用到系统..."
    
    local install_dir="/opt/$APP_NAME"
    local bin_dir="/usr/local/bin"
    
    # 创建安装目录
    sudo mkdir -p "$install_dir"
    sudo mkdir -p "$install_dir/logs"
    sudo mkdir -p "$install_dir/config"
    
    # 复制文件
    sudo cp -r "$BUILD_DIR"/* "$install_dir/"
    sudo chown -R $USER_NAME:$USER_NAME "$install_dir"
    sudo chmod +x "$install_dir/$APP_NAME"
    
    # 创建符号链接
    sudo ln -sf "$install_dir/$APP_NAME" "$bin_dir/$APP_NAME"
    
    log_success "应用安装完成"
}

# 创建systemd服务
create_service() {
    log_info "创建systemd服务..."
    
    local service_file="/etc/systemd/system/${SERVICE_NAME}.service"
    
    sudo tee "$service_file" > /dev/null << EOF
[Unit]
Description=Excel to Word Conversion Web Server
After=network.target
Wants=network.target

[Service]
Type=simple
User=$USER_NAME
Group=$USER_NAME
WorkingDirectory=/opt/$APP_NAME
ExecStart=/opt/$APP_NAME/$APP_NAME
Restart=always
RestartSec=10
StandardOutput=journal
StandardError=journal
SyslogIdentifier=$SERVICE_NAME

# 环境变量
Environment=RUST_LOG=info
Environment=PORT=$PORT

# 安全设置
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=strict
ProtectHome=true
ReadWritePaths=/opt/$APP_NAME/logs

[Install]
WantedBy=multi-user.target
EOF
    
    # 重新加载systemd
    sudo systemctl daemon-reload
    sudo systemctl enable $SERVICE_NAME
    
    log_success "systemd服务创建完成"
}

# 配置防火墙
configure_firewall() {
    log_info "配置防火墙..."
    
    # 检查防火墙类型
    if command -v ufw &> /dev/null; then
        # Ubuntu/Debian UFW
        sudo ufw allow $PORT/tcp
        log_success "UFW防火墙规则已添加"
    elif command -v firewall-cmd &> /dev/null; then
        # CentOS/RHEL firewalld
        sudo firewall-cmd --permanent --add-port=$PORT/tcp
        sudo firewall-cmd --reload
        log_success "firewalld防火墙规则已添加"
    else
        log_warning "未检测到支持的防火墙，请手动开放端口 $PORT"
    fi
}

# 创建Nginx反向代理配置（可选）
create_nginx_config() {
    log_info "创建Nginx反向代理配置..."
    
    local nginx_conf="/etc/nginx/sites-available/$SERVICE_NAME"
    
    if command -v nginx &> /dev/null; then
        sudo tee "$nginx_conf" > /dev/null << EOF
server {
    listen 80;
    server_name _;  # 替换为你的域名
    
    location / {
        proxy_pass http://127.0.0.1:$PORT;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        
        # 文件上传大小限制
        client_max_body_size 50M;
        proxy_read_timeout 300s;
        proxy_connect_timeout 75s;
    }
}
EOF
        
        # 启用站点
        sudo ln -sf "$nginx_conf" "/etc/nginx/sites-enabled/"
        sudo nginx -t && sudo systemctl reload nginx
        
        log_success "Nginx配置创建完成"
    else
        log_warning "未检测到Nginx，跳过反向代理配置"
    fi
}

# 启动服务
start_service() {
    log_info "启动服务..."
    
    sudo systemctl start $SERVICE_NAME
    sudo systemctl status $SERVICE_NAME --no-pager
    
    if sudo systemctl is-active --quiet $SERVICE_NAME; then
        log_success "服务启动成功"
        log_info "服务状态: $(sudo systemctl is-active $SERVICE_NAME)"
        log_info "服务地址: http://$(hostname -I | awk '{print $1}'):$PORT"
    else
        log_error "服务启动失败"
        sudo journalctl -u $SERVICE_NAME -n 20
        exit 1
    fi
}

# 健康检查
health_check() {
    log_info "执行健康检查..."
    
    sleep 5  # 等待服务完全启动
    
    if curl -s "http://localhost:$PORT" > /dev/null; then
        log_success "健康检查通过"
    else
        log_error "健康检查失败"
        sudo journalctl -u $SERVICE_NAME -n 20
        exit 1
    fi
}

# 显示部署信息
show_deploy_info() {
    log_success "部署完成！"
    echo
    echo "=== 部署信息 ==="
    echo "应用名称: $APP_NAME"
    echo "服务名称: $SERVICE_NAME"
    echo "安装目录: /opt/$APP_NAME"
    echo "服务状态: $(sudo systemctl is-active $SERVICE_NAME)"
    echo "服务地址: http://$(hostname -I | awk '{print $1}'):$PORT"
    echo
    echo "=== 常用命令 ==="
    echo "查看服务状态: sudo systemctl status $SERVICE_NAME"
    echo "查看服务日志: sudo journalctl -u $SERVICE_NAME -f"
    echo "重启服务: sudo systemctl restart $SERVICE_NAME"
    echo "停止服务: sudo systemctl stop $SERVICE_NAME"
    echo
    echo "=== 文件位置 ==="
    echo "二进制文件: /opt/$APP_NAME/$APP_NAME"
    echo "配置文件: /opt/$APP_NAME/config/"
    echo "日志目录: /opt/$APP_NAME/logs/"
    echo "systemd服务: /etc/systemd/system/${SERVICE_NAME}.service"
}

# 主函数
main() {
    echo "=========================================="
    echo "    Excel到Word转换器部署脚本"
    echo "=========================================="
    echo
    
    check_root
    check_dependencies
    build_app
    create_user
    install_app
    create_service
    configure_firewall
    create_nginx_config
    start_service
    health_check
    show_deploy_info
}

# 运行主函数
main "$@"
