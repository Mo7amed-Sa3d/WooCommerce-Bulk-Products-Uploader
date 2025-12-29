import queue
import threading
import time
from typing import Dict, Any
import logging

logger = logging.getLogger(__name__)

class UploadQueueManager:
    def __init__(self, wc_api, wp_api, max_workers=3):
        self.wc_api = wc_api
        self.wp_api = wp_api
        self.upload_queue = queue.Queue()
        self.results_queue = queue.Queue()
        self.running = True
        self.max_workers = max_workers
        self.workers = []
        
        # Stats
        self.stats = {
            'completed': 0,
            'failed': 0,
            'total': 0
        }
        
        # Start worker threads
        for i in range(max_workers):
            worker = threading.Thread(
                target=self._process_queue_worker,
                args=(f"Worker-{i+1}",),
                daemon=True
            )
            worker.start()
            self.workers.append(worker)
        
        # Start results processor
        self.results_thread = threading.Thread(
            target=self._process_results,
            daemon=True
        )
        self.results_thread.start()
    
    def _process_queue_worker(self, worker_name):
        """Worker thread function to process upload tasks"""
        while self.running:
            try:
                # Get task from queue (wait up to 1 second)
                task = self.upload_queue.get(timeout=1)
                logger.info(f"{worker_name} processing: {task.get('title', 'Unknown')}")
                
                # Upload images first
                uploaded_images = []
                for i, image_path in enumerate(task['images']):
                    result = self.wp_api.upload_media(image_path)
                    if result['success']:
                        uploaded_images.append(result)
                        logger.info(f"{worker_name}: Uploaded image {i+1}/{len(task['images'])}")
                    else:
                        logger.error(f"{worker_name}: Failed to upload {image_path}")
                        # Continue with other images even if one fails
                
                if not uploaded_images:
                    self.results_queue.put({
                        'success': False,
                        'title': task['title'],
                        'error': 'No images uploaded successfully',
                        'task': task
                    })
                    self.upload_queue.task_done()
                    continue
                
                # Prepare product data for WooCommerce
                wc_product_data = {
                    'name': task['title'],
                    'description': task['description'],
                    'type': 'simple',
                    'regular_price': task['price'],
                    'categories': [{'id': task['category_id']}],
                    'sku': task.get('sku', ''),
                    'images': []
                }
                
                # Add images to product
                for i, img_info in enumerate(uploaded_images):
                    image_data = {'id': img_info['id']}
                    if i == 0:  # First image is featured
                        wc_product_data['images'].insert(0, image_data)
                    else:
                        wc_product_data['images'].append(image_data)
                
                # Create product in WooCommerce
                result = self.wc_api.create_product(wc_product_data)
                result['task'] = task
                
                # Update stats
                if result['success']:
                    self.stats['completed'] += 1
                else:
                    self.stats['failed'] += 1
                
                # Put result in results queue
                self.results_queue.put(result)
                
                # Mark task as done
                self.upload_queue.task_done()
                logger.info(f"{worker_name} completed: {task.get('title', 'Unknown')}")
                
            except queue.Empty:
                continue
            except Exception as e:
                logger.error(f"{worker_name} error: {e}")
                self.stats['failed'] += 1
                self.results_queue.put({
                    'success': False,
                    'title': task.get('title', 'Unknown'),
                    'error': str(e),
                    'task': task
                })
                self.upload_queue.task_done()
    
    def _process_results(self):
        """Process results from uploads (can be overridden for GUI updates)"""
        while self.running:
            try:
                result = self.results_queue.get(timeout=1)
                if hasattr(self, 'on_upload_complete'):
                    self.on_upload_complete(result)
                self.results_queue.task_done()
            except queue.Empty:
                continue
            except Exception as e:
                logger.error(f"Results processor error: {e}")
    
    def on_upload_complete(self, result):
        """Callback for upload completion (override in GUI)"""
        # This should be overridden by the GUI
        pass
    
    def add_to_queue(self, product_data: Dict[str, Any]):
        """Add a product to the upload queue"""
        self.upload_queue.put(product_data)
        self.stats['total'] += 1
        return self.upload_queue.qsize()
    
    def get_queue_size(self):
        """Get current queue size"""
        return self.upload_queue.qsize()
    
    def get_active_workers(self):
        """Get number of active worker threads"""
        return sum(1 for w in self.workers if w.is_alive())
    
    def get_stats(self):
        """Get upload statistics"""
        return self.stats.copy()
    
    def stop(self):
        """Stop all workers"""
        self.running = False
        for worker in self.workers:
            if worker.is_alive():
                worker.join(timeout=2)
        if self.results_thread.is_alive():
            self.results_thread.join(timeout=2)
    
    def wait_for_completion(self, timeout=None):
        """Wait for all tasks to complete"""
        self.upload_queue.join()